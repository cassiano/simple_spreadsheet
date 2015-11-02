require_relative '../simple_spreadsheet'
require 'contest'
require 'turn/autorun'

class TestSpreadsheet < Test::Unit::TestCase
  context 'SpreadSheet' do
    setup do
      @spreadsheet = Spreadsheet.new
    end

    teardown do
      assert @spreadsheet.consistent? unless @skip_teardown
    end

    test 'new and empty spreadsheets have no cells' do
      assert_equal 0, @spreadsheet.cell_count
    end

    test '#find_or_create_cell finds preexistent cells' do
      a1 = @spreadsheet.set :A1

      assert_equal a1, @spreadsheet.find_or_create_cell(:A1)
    end

    test '#find_or_create_cell creates new cells as needed' do
      a1 = @spreadsheet.find_or_create_cell(:A1)

      assert_equal :A1, a1.ref.ref
    end

    test 'cells can be found using symbol or string case insensitive references' do
      a1 = @spreadsheet.set :A1

      assert_equal a1, @spreadsheet.find_or_create_cell(:A1)
      assert_equal a1, @spreadsheet.find_or_create_cell(:a1)
      assert_equal a1, @spreadsheet.find_or_create_cell('A1')
      assert_equal a1, @spreadsheet.find_or_create_cell('a1')
    end

    # test '#add_column' do
    #   a1 = @spreadsheet.set :A1, 1
    #   b1 = @spreadsheet.set :B1, 2
    #   a2 = @spreadsheet.set :A2, 3
    #   b2 = @spreadsheet.set :B2, 4
    #
    #   @spreadsheet.add_column :A
    #
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:A1)
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:A2)
    #   assert_equal 1,                                 @spreadsheet.find_or_create_cell(:B1)
    #   assert_equal 2,                                 @spreadsheet.find_or_create_cell(:B2)
    #   assert_equal 3,                                 @spreadsheet.find_or_create_cell(:C1)
    #   assert_equal 4,                                 @spreadsheet.find_or_create_cell(:C2)
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:D1)
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:D2)
    # end

    # test '#add_row' do
    #   a1 = @spreadsheet.set :A1, 1
    #   b1 = @spreadsheet.set :B1, 2
    #   a2 = @spreadsheet.set :A2, 3
    #   b2 = @spreadsheet.set :B2, 4
    #
    #   @spreadsheet.add_row 1
    #
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:A1)
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:B1)
    #   assert_equal 1,                                 @spreadsheet.find_or_create_cell(:A2)
    #   assert_equal 2,                                 @spreadsheet.find_or_create_cell(:B2)
    #   assert_equal 3,                                 @spreadsheet.find_or_create_cell(:A3)
    #   assert_equal 4,                                 @spreadsheet.find_or_create_cell(:B3)
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:A4)
    #   assert_equal Cell::DEFAULT_VALUE,  @spreadsheet.find_or_create_cell(:B4)
    # end

    context 'Cell' do
      test 'can hold scalar values, like numbers and strings' do
        a1 = @spreadsheet.set :A1, 1
        a2 = @spreadsheet.set :A2, 1.234
        a3 = @spreadsheet.set :A3, 'foo bar'

        assert_equal 1, a1.eval
        assert_equal 1.234, a2.eval
        assert_equal 'foo bar', a3.eval
      end

      test 'can hold formulas, referencing other cells' do
        a1 = @spreadsheet.set :A1, 1
        a2 = @spreadsheet.set :A2, 2
        a3 = @spreadsheet.set :A3, '= (A1 + A2) * 3'

        assert_equal 1, a1.eval
        assert_equal 2, a2.eval
        assert_equal (1 + 2) * 3, a3.eval
      end

      test 'cannot have circular references in formulas' do
        @skip_teardown = true

        assert_nothing_raised do
          a1 = @spreadsheet.set :A1, '= A2'
          a2 = @spreadsheet.set :A2, '= A3'
          a3 = @spreadsheet.set :A3, '= A4'
          a4 = @spreadsheet.set :A4, '= A5'
        end

        assert_raises Cell::CircularReferenceError do
          a5 = @spreadsheet.set :A5, '= A1'
        end
      end

      test 'cannot have auto references in formulas' do
        @skip_teardown = true

        assert_raises Cell::CircularReferenceError do
          a1 = @spreadsheet.set :A1, '= A1'
        end
      end

      test 'have default values' do
        a1 = @spreadsheet.set :A1

        assert_equal Cell::DEFAULT_VALUE, a1.eval
      end

      test 'have empty references and observers when created' do
        a1 = @spreadsheet.set :A1

        assert_equal 0, a1.references.count
        assert_equal 0, a1.observers.count
      end

      test 'saves references to other cells' do
        a1 = @spreadsheet.set :A1, '= A2 + A3 + A4'
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
        a1 = @spreadsheet.set :A1, '= A2 + A3 + A4'
        a2 = @spreadsheet.find_or_create_cell :A2
        a3 = @spreadsheet.find_or_create_cell :A3
        a4 = @spreadsheet.find_or_create_cell :A4

        observers = Set.new
        observers << a1

        assert_equal observers, a2.observers
        assert_equal observers, a3.observers
        assert_equal observers, a4.observers
      end

      test 'can hold formulas with buitin functions' do
        a1 = @spreadsheet.set :A1, 1
        a2 = @spreadsheet.set :A2, 2
        a3 = @spreadsheet.set :A3, 4
        a4 = @spreadsheet.set :A4, '= sum(A1, A2, A3)'

        assert_equal 1 + 2 + 4, a4.eval
      end

      test 'can hold formulas with any Ruby code' do
        a1 = @spreadsheet.set :A1, 'foo'
        a2 = @spreadsheet.set :A2, 'bar'
        a3 = @spreadsheet.set :A3, '= "A1" + " " + "A2"'

        assert_equal 'foo bar', a3.eval
      end

      test '.splat_range works for cells in same row' do
        assert_equal [[:A1, :B1, :C1]], CellRef.splat_range(CellRef.new(:A1), CellRef.new(:C1))
      end

      test '.splat_range works for cells in same column' do
        assert_equal [[:A1], [:A2], [:A3]], CellRef.splat_range(CellRef.new(:A1), CellRef.new(:A3))
      end

      test '.splat_range works for cells in distinct rows and columns' do
        assert_equal [[:A1, :B1, :C1], [:A2, :B2, :C2], [:A3, :B3, :C3]], CellRef.splat_range(CellRef.new(:A1), CellRef.new(:C3))
      end

      test 'can hold formulas with functions which include cell ranges' do
        a1 = @spreadsheet.set :A1, 1
        b1 = @spreadsheet.set :B1, 2
        c1 = @spreadsheet.set :C1, 4
        a2 = @spreadsheet.set :A2, 8
        b2 = @spreadsheet.set :B2, 16
        c2 = @spreadsheet.set :C2, 32
        a3 = @spreadsheet.set :A3, 64
        b3 = @spreadsheet.set :B3, 128
        c4 = @spreadsheet.set :C3, 256
        a4 = @spreadsheet.set :A4, '= sum(A1:C3)'

        assert_equal 1 + 2 + 4 + 8 + 16 + 32 + 64 + 128 + 256, a4.eval
      end

      test 'changes in references are automatically reflected in dependent cells' do
        a1 = @spreadsheet.set :A1, 1
        a2 = @spreadsheet.set :A2, 2
        a3 = @spreadsheet.set :A3, '= (A1 + A2) * 3'

        a1.content = 10
        a2.content = 20
        assert_equal (10 + 20) * 3, a3.eval
      end

      test 'changes in references are automatically reflected in dependent cells, even if within a range' do
        a1 = @spreadsheet.set :A1, 1
        a2 = @spreadsheet.set :A2, 2
        a3 = @spreadsheet.set :A3, 4
        a4 = @spreadsheet.set :A4, '= sum(A1:A3)'

        assert_equal 1 + 2 + 4, a4.eval

        a2.content = 8
        assert_equal 1 + 8 + 4, a4.eval
      end

      test '#move_to!' do
        old_a1 = @spreadsheet.set :A1, 1
        a2     = @spreadsheet.set :A2, 2
        a3     = @spreadsheet.set :A3, '= A1 + A2'
        a4     = @spreadsheet.set :A4, '= A3'

        assert_equal (a3_value = 1 + 2), a3.eval

        a3_last_evaluated_at = a3.last_evaluated_at
        a4_last_evaluated_at = a4.last_evaluated_at

        a1_observers = Set.new
        a1_observers << a3
        assert_equal a1_observers, old_a1.observers

        # Move A1 to C5, so (old) A1 actually "becomes" C5.
        old_a1.move_to! :C5

        # Assert A3's formula and references have been updated and that it's (evaluated) value hasn't changed.
        c5 = @spreadsheet.find_or_create_cell :C5
        a3_references = Set.new
        a3_references << c5
        a3_references << a2
        assert_equal old_a1, c5
        assert_equal '= C5 + A2', a3.content
        assert_equal a3_references, a3.references
        assert_equal a3_value, a3.eval
        assert_not_equal a3_last_evaluated_at, a3.last_evaluated_at

        # Assert (new) A1 cell is empty and has no (more) observers.
        new_a1           = @spreadsheet.find_or_create_cell :A1
        new_a1_observers = Set.new
        assert_equal Cell::DEFAULT_VALUE, new_a1.eval
        assert_equal new_a1_observers, new_a1.observers

        # Assert C5 is now being observed by A3, instead of A1.
        c5_observers = Set.new
        c5_observers << a3
        assert_equal c5_observers, c5.observers

        # Assert A4 hasn't been reevaluated, since A3's value never actually changed.
        assert_equal a4_last_evaluated_at, a4.last_evaluated_at
      end

      test '#copy_to' do
        a1 = @spreadsheet.set :A1, 1
        a2 = @spreadsheet.set :A2, 2
        a3 = @spreadsheet.set :A3, '= A1 + A2'
        b1 = @spreadsheet.set :B1, 10
        b2 = @spreadsheet.set :B2, 20

        assert_equal (a3_value = 1 + 2), a3.eval

        # Copy A3 to B3.
        a3.copy_to :B3

        # Assert A3's formula and references have been updated and that it's (evaluated) value hasn't changed.
        b3 = @spreadsheet.find_or_create_cell :B3
        b3_references = Set.new
        b3_references << b1
        b3_references << b2
        assert_equal '= B1 + B2', b3.content
        assert_equal 10 + 20, b3.eval
      end

      test '#move_right!' do
        a1 = @spreadsheet.set :A1, 1

        a1.move_right!

        new_a1 = @spreadsheet.find_or_create_cell :A1
        b1     = @spreadsheet.find_or_create_cell :B1

        assert_equal a1, b1
        assert_equal 1, b1.eval
        assert_equal Cell::DEFAULT_VALUE, new_a1.eval
      end

      test '#move_right! should allow to move more that 1 column (default value)' do
        a1 = @spreadsheet.set :A1, 1

        a1.move_right! 4

        new_a1 = @spreadsheet.find_or_create_cell :A1
        e1     = @spreadsheet.find_or_create_cell :E1

        assert_equal a1, e1
        assert_equal 1, e1.eval
        assert_equal Cell::DEFAULT_VALUE, new_a1.eval
      end

      test '#move_left!' do
        b1 = @spreadsheet.set :B1, 1

        b1.move_left!

        new_b1 = @spreadsheet.find_or_create_cell :B1
        a1     = @spreadsheet.find_or_create_cell :A1

        assert_equal a1, b1
        assert_equal 1, b1.eval
        assert_equal Cell::DEFAULT_VALUE, new_b1.eval
      end

      test '#move_left! should allow to move more that 1 column (default value)' do
        e1 = @spreadsheet.set :E1, 1

        e1.move_left! 4

        new_e1 = @spreadsheet.find_or_create_cell :E1
        a1     = @spreadsheet.find_or_create_cell :A1

        assert_equal e1, a1
        assert_equal 1, a1.eval
        assert_equal Cell::DEFAULT_VALUE, new_e1.eval
      end

      test '#move_left! should raise an error when in leftmost cell' do
        a1 = @spreadsheet.set :A1

        assert_raises CellRef::IllegalCellReference do
          a1.move_left!
        end
      end

      test '#move_down!' do
        a1 = @spreadsheet.set :A1, 1

        a1.move_down!

        new_a1 = @spreadsheet.find_or_create_cell :A1
        a2     = @spreadsheet.find_or_create_cell :A2

        assert_equal a1, a2
        assert_equal 1, a2.eval
        assert_equal Cell::DEFAULT_VALUE, new_a1.eval
      end

      test '#move_down! should allow to move more that 1 row (default value)' do
        a1 = @spreadsheet.set :A1, 1

        a1.move_down! 4

        new_a1 = @spreadsheet.find_or_create_cell :A1
        a5     = @spreadsheet.find_or_create_cell :A5

        assert_equal a1, a5
        assert_equal 1, a5.eval
        assert_equal Cell::DEFAULT_VALUE, new_a1.eval
      end

      test '#move_up!' do
        a2 = @spreadsheet.set :A2, 1

        a2.move_up!

        new_a2 = @spreadsheet.find_or_create_cell :A2
        a1     = @spreadsheet.find_or_create_cell :A1

        assert_equal a2, a1
        assert_equal 1, a1.eval
        assert_equal Cell::DEFAULT_VALUE, new_a2.eval
      end

      test '#move_up! should allow to move more that 1 row (default value)' do
        a5 = @spreadsheet.set :A5, 1

        a5.move_up! 4

        new_a5 = @spreadsheet.find_or_create_cell :A5
        a1     = @spreadsheet.find_or_create_cell :A1

        assert_equal a5, a1
        assert_equal 1, a1.eval
        assert_equal Cell::DEFAULT_VALUE, new_a5.eval
      end

      test '#move_up! should raise an error when in topmost cell' do
        a1 = @spreadsheet.set :A1

        assert_raises CellRef::IllegalCellReference do
          a1.move_up!
        end
      end
    end
  end
end
