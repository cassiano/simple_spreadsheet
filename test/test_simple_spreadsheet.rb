require_relative '../simple_spreadsheet'
require 'contest'
require 'turn/autorun'

class TestSpreadsheet < Test::Unit::TestCase
  context 'CellRef' do
    test '#initialize can receive symbols, strings or even arrays, all case insensitive' do
      [
        CellRef.new(:A2),
        CellRef.new(:a2),
        CellRef.new('A2'),
        CellRef.new('a2'),
        CellRef.new('A', 2),
        CellRef.new('a', 2),
        CellRef.new(:A, 2),
        CellRef.new(:a, 2),
        CellRef.new(1, 2),
        CellRef.new(['A', 2]),
        CellRef.new(['a', 2]),
        CellRef.new([:A, 2]),
        CellRef.new([:a, 2]),
        CellRef.new([1, 2])
      ].each do |ref|
        assert_equal :A2, ref.ref
      end

      [
        :A,
        2,
        '2A',
        'A 2',
        nil
      ].each do |invalid_ref|
        assert_raises CellRef::IllegalCellReference do
          CellRef.new invalid_ref
        end
      end
    end

    test '#col' do
      ref_a1   = CellRef.new(:A1)
      ref_aa1  = CellRef.new(:AA1)
      ref_aaa1 = CellRef.new(:AAA1)

      assert_equal :A,    ref_a1.col
      assert_equal :AA,   ref_aa1.col
      assert_equal :AAA,  ref_aaa1.col
    end

    test '#col_index' do
      ref_a1 = CellRef.new(:A1)

      assert_equal 1, ref_a1.col_index
    end

    test '#row' do
      ref_a2 = CellRef.new(:A2)

      assert_equal 2, ref_a2.row
    end

    test '#col_and_row' do
      ref_a2 = CellRef.new(:A2)

      assert_equal [:A, 2], ref_a2.col_and_row
    end

    test '#==' do
      ref_a1         = CellRef.new(:A1)
      ref_another_a1 = CellRef.new(:A1)

      assert_equal ref_another_a1, ref_a1
    end

    test '#neighbor' do
      ref_d5 = CellRef.new(:D5)

      assert_equal ref_d5, ref_d5.neighbor
      assert_equal CellRef.new(:G5), ref_d5.neighbor(col_count: 3)
      assert_equal CellRef.new(:A5), ref_d5.neighbor(col_count: -3)
      assert_raises CellRef::IllegalCellReference do
        ref_d5.neighbor(col_count: -4)
      end

      assert_equal CellRef.new(:D9), ref_d5.neighbor(row_count: 4)
      assert_equal CellRef.new(:D1), ref_d5.neighbor(row_count: -4)
      assert_raises CellRef::IllegalCellReference do
        ref_d5.neighbor(row_count: -5)
      end

      assert_equal CellRef.new(:G1), ref_d5.neighbor(col_count: 3, row_count: -4)
      assert_raises CellRef::IllegalCellReference do
        ref_d5.neighbor(col_count: -4, row_count: -5)
      end
    end

    test '#left_neighbor' do
      ref_d5 = CellRef.new(:D5)

      assert_equal CellRef.new(:C5), ref_d5.left_neighbor
      assert_equal CellRef.new(:A5), ref_d5.left_neighbor(3)
      assert_raises CellRef::IllegalCellReference do
        ref_d5.left_neighbor(4)
      end
    end

    test '#right_neighbor' do
      ref_d5 = CellRef.new(:D5)

      assert_equal CellRef.new(:E5), ref_d5.right_neighbor
      assert_equal CellRef.new(:G5), ref_d5.right_neighbor(3)
    end

    test '#upper_neighbor' do
      ref_d5 = CellRef.new(:D5)

      assert_equal CellRef.new(:D4), ref_d5.upper_neighbor
      assert_equal CellRef.new(:D1), ref_d5.upper_neighbor(4)
      assert_raises CellRef::IllegalCellReference do
        ref_d5.upper_neighbor(5)
      end
    end

    test '#lower_neighbor' do
      ref_d5 = CellRef.new(:D5)

      assert_equal CellRef.new(:D6), ref_d5.lower_neighbor
      assert_equal CellRef.new(:D9), ref_d5.lower_neighbor(4)
    end

    test '#to_s' do
      ref_d5 = CellRef.new(:D5)

      assert_equal 'D5', ref_d5.to_s
    end
  end

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
        a1_ref = CellRef.new(:A1)
        b1_ref = CellRef.new(:B1)
        c1_ref = CellRef.new(:C1)

        assert_equal [[a1_ref, b1_ref, c1_ref]], CellRef.splat_range(:A1, :C1)
        assert_equal [[a1_ref, b1_ref, c1_ref]], CellRef.splat_range(CellRef.new(:A1), CellRef.new(:C1))
      end

      test '.splat_range works for cells in same column' do
        a1_ref = CellRef.new(:A1)
        a2_ref = CellRef.new(:A2)
        a3_ref = CellRef.new(:A3)

        assert_equal [[a1_ref], [a2_ref], [a3_ref]], CellRef.splat_range(:A1, :A3)
        assert_equal [[a1_ref], [a2_ref], [a3_ref]], CellRef.splat_range(CellRef.new(:A1), CellRef.new(:A3))
      end

      test '.splat_range works for cells in distinct rows and columns' do
        a1_ref = CellRef.new(:A1)
        b1_ref = CellRef.new(:B1)
        c1_ref = CellRef.new(:C1)
        a2_ref = CellRef.new(:A2)
        b2_ref = CellRef.new(:B2)
        c2_ref = CellRef.new(:C2)
        a3_ref = CellRef.new(:A3)
        b3_ref = CellRef.new(:B3)
        c3_ref = CellRef.new(:C3)

        assert_equal [
          [a1_ref, b1_ref, c1_ref],
          [a2_ref, b2_ref, c2_ref],
          [a3_ref, b3_ref, c3_ref]
        ], CellRef.splat_range(:A1, :C3)

        assert_equal [
          [a1_ref, b1_ref, c1_ref],
          [a2_ref, b2_ref, c2_ref],
          [a3_ref, b3_ref, c3_ref]
        ], CellRef.splat_range(CellRef.new(:A1), CellRef.new(:C3))
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
