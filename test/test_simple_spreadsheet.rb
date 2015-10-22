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
  end
end
