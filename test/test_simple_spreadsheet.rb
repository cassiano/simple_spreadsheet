require_relative '../simple_spreadsheet'
require 'contest'
require 'turn/autorun'

class TestSpreadsheet < Test::Unit::TestCase
  context 'SpreadSheet' do
    setup do
      @spreadsheet = Spreadsheet.new
    end

    test 'new and empty spreadsheets have no cells' do
      assert_equal 0, @spreadsheet.cells.count
    end

    test '#find_or_create_cell finds preexistent cells' do
      a1 = @spreadsheet.add_cell :A1

      assert_equal a1, @spreadsheet.find_or_create_cell(:A1)
    end

    test '#find_or_create_cell creates new cells as needed' do
      a1 = @spreadsheet.find_or_create_cell(:A1)

      assert_equal :A1, a1.ref
    end

    test 'cells can be found using symbol or string case insensitive references' do
      a1 = @spreadsheet.add_cell :A1

      assert_equal a1, @spreadsheet.find_or_create_cell(:A1)
      assert_equal a1, @spreadsheet.find_or_create_cell(:a1)
      assert_equal a1, @spreadsheet.find_or_create_cell('A1')
      assert_equal a1, @spreadsheet.find_or_create_cell('a1')
    end

    context 'Cell' do
      test 'can hold scalar values, like numbers and strings' do
        contents = {
          a1: 1,
          a2: 1.234,
          a3: 'Some Text'
        }

        a1 = @spreadsheet.add_cell :A1, contents[:a1]
        a2 = @spreadsheet.add_cell :A2, contents[:a2]
        a3 = @spreadsheet.add_cell :A3, contents[:a3]

        assert_equal contents[:a1], a1.eval
        assert_equal contents[:a2], a2.eval
        assert_equal contents[:a3], a3.eval
      end

      test 'can hold formulas, referencing other cells' do
        contents = {
          a1: 1,
          a2: 2,
          a3: '= (A1 + A2) * 3'
        }

        a1 = @spreadsheet.add_cell :A1, contents[:a1]
        a2 = @spreadsheet.add_cell :A2, contents[:a2]
        a3 = @spreadsheet.add_cell :A3, contents[:a3]

        assert_equal contents[:a1], a1.eval
        assert_equal contents[:a2], a2.eval
        assert_equal (contents[:a1] + contents[:a2]) * 3, a3.eval
      end

      test 'cannot have circular references in formulas' do
        contents = {
          a1: '= A2',
          a2: '= A3',
          a3: '= A4',
          a4: '= A5',
          a5: '= A1'
        }

        assert_nothing_raised do
          a1 = @spreadsheet.find_or_create_cell :A1, contents[:a1]
          a2 = @spreadsheet.find_or_create_cell :A2, contents[:a2]
          a3 = @spreadsheet.find_or_create_cell :A3, contents[:a3]
          a4 = @spreadsheet.find_or_create_cell :A4, contents[:a4]
        end

        assert_raises Spreadsheet::Cell::CircularReferenceError do
          a5 = @spreadsheet.find_or_create_cell :A5, contents[:a5]
        end
      end

      test 'have default values' do
        a1 = @spreadsheet.add_cell :A1

        assert_equal Spreadsheet::Cell::DEFAULT_VALUE, a1.eval
      end

      test 'have empty references and observers when created' do
        a1 = @spreadsheet.add_cell :A1

        assert_equal 0, a1.references.count
        assert_equal 0, a1.observers.count
      end

      test 'saves references to other cells' do
        a1 = @spreadsheet.add_cell :A1, '= A2 + A3 + A4'
        a2 = @spreadsheet.find_or_create_cell :A2
        a3 = @spreadsheet.find_or_create_cell :A3
        a4 = @spreadsheet.find_or_create_cell :A4

        references = Set.new
        references << a2
        references << a3
        references << a4

        assert_equal references, a1.references
      end

      test 'are marked as observers in other cells when referencing them' do
        a1 = @spreadsheet.add_cell :A1, '= A2 + A3 + A4'
        a2 = @spreadsheet.find_or_create_cell :A2
        a3 = @spreadsheet.find_or_create_cell :A3
        a4 = @spreadsheet.find_or_create_cell :A4

        observers = Set.new
        observers << a1

        assert_equal observers, a2.observers
        assert_equal observers, a3.observers
        assert_equal observers, a4.observers
      end
    end

    test 'can hold formulas with functions' do
      contents = {
        a1: 1,
        a2: 2,
        a3: 3,
        a4: '= sum(A1, A2, A3)',
        a5: '= sum([A1, A2, A3])'
      }

      a1 = @spreadsheet.add_cell :A1, contents[:a1]
      a2 = @spreadsheet.add_cell :A2, contents[:a2]
      a3 = @spreadsheet.add_cell :A3, contents[:a3]
      a4 = @spreadsheet.add_cell :A4, contents[:a4]
      a5 = @spreadsheet.add_cell :A5, contents[:a5]

      assert_equal Spreadsheet::Formula.sum(contents[:a1], contents[:a2], contents[:a3]), a4.eval
      assert_equal Spreadsheet::Formula.sum(contents[:a1], contents[:a2], contents[:a3]), a5.eval
    end

    test '.cells_in_range works for cells in same row' do
      assert_equal [[:A1, :B1, :C1]], Spreadsheet::Cell.cells_in_range(:A1, :C1)
    end

    test '.cells_in_range works for cells in same column' do
      assert_equal [[:A1], [:A2], [:A3]], Spreadsheet::Cell.cells_in_range(:A1, :A3)
    end

    test '.cells_in_range works for cells in distinct rows and columns' do
      assert_equal [[:A1, :B1, :C1], [:A2, :B2, :C2], [:A3, :B3, :C3]], Spreadsheet::Cell.cells_in_range(:A1, :C3)
    end

    test 'can hold formulas with functions which include cell ranges' do
      contents = {
        a1: 1,
        b1: 2,
        c1: 3,
        a2: 1,
        b2: 2,
        c2: 3,
        a3: 1,
        b3: 2,
        c3: 3,
        a4: '= sum(A1:C3)'
      }

      a1 = @spreadsheet.add_cell :A1, contents[:a1]
      b1 = @spreadsheet.add_cell :B1, contents[:b1]
      c1 = @spreadsheet.add_cell :C1, contents[:c1]
      a2 = @spreadsheet.add_cell :A2, contents[:a2]
      b2 = @spreadsheet.add_cell :B2, contents[:b2]
      c2 = @spreadsheet.add_cell :C2, contents[:c2]
      a3 = @spreadsheet.add_cell :A3, contents[:a2]
      b3 = @spreadsheet.add_cell :B3, contents[:b2]
      c4 = @spreadsheet.add_cell :C3, contents[:c2]
      a4 = @spreadsheet.add_cell :A4, contents[:a4]

      assert_equal Spreadsheet::Formula.sum(
        contents[:a1], contents[:b1], contents[:c1],
        contents[:a2], contents[:b2], contents[:c2],
        contents[:a3], contents[:b3], contents[:c3]
      ), a4.eval
    end

    test 'changes in references are automatically reflected in dependent cells' do
    end
  end
end
