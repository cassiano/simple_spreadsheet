DEBUG = false

require 'set'

class Spreadsheet
  # List of possible exceptions.
  class AlreadyExistentCellError < StandardError; end

  attr_reader :cells

  def initialize
    @cells = {}
  end

  def update_cell_ref(ref, cell)
    cells[ref] = cell
  end

  def find_or_create_cell(ref, content = nil)
    ref = ref.upcase.to_sym

    (cells[ref] || add_cell(ref)).tap do |cell|
      cell.content = content if content
    end
  end

  alias_method :set, :find_or_create_cell

  def add_cell(ref, content = nil)
    raise AlreadyExistentCellError, "Cell #{ref} already exists" if cells[ref]

    ref = ref.upcase.to_sym

    Cell.new(self, ref, content).tap do |cell|
      cells[ref] = cell
      # update_cell_ref ref, cell
    end
  end

  def update_cell_ref(old_ref, new_ref)
    cells[new_ref] = cells.delete(old_ref)
  end

  class Cell
    CELL_REF1          = '\b[A-Z]+[1-9]\d*\b'
    CELL_REF2          = '\b([A-Z]+)([1-9]\d*)\b'
    CELL_REF_REG_EXP   = /#{CELL_REF1}/i
    CELL_REF2_REG_EXP  = /#{CELL_REF2}/i
    CELL_RANGE_REG_EXP = /((#{CELL_REF1}):(#{CELL_REF1}))/i
    DEFAULT_VALUE      = 0
    COL_RANGE          = ('A'..'ZZZ').to_a.map(&:to_sym)

    # List of possible exceptions.
    class CircularReferenceError < StandardError; end

    attr_reader :spreadsheet, :ref, :references, :observers, :content, :last_evaluated_at

    def initialize(spreadsheet, ref, content = nil)
      puts "Creating cell #{ref}" if DEBUG

      ref = ref.upcase.to_sym

      @spreadsheet = spreadsheet
      @ref         = ref
      @references  = Set.new
      @observers   = Set.new

      self.content = content
    end

    def add_observer(cell)
      puts "Adding observer #{cell.ref} to #{ref}" if DEBUG

      observers << cell
    end

    def remove_observer(cell)
      puts "Removing observer #{cell.ref} from #{ref}" if DEBUG

      observers.delete cell
    end

    def content=(new_content)
      new_content = new_content.strip if String === new_content

      old_references = Set.new
      new_references = Set.new

      old_references = references.clone if is_formula?

      @content = new_content

      if is_formula?
        # Splat ranges, e.g., replace 'A1:A3' by '[[A1, A2, A3]]'.
        new_content[1..-1].scan(CELL_RANGE_REG_EXP).each do |(range, upper_left_ref, lower_right_ref)|
          new_content.gsub! range, self.class.splat_range(upper_left_ref, lower_right_ref).to_s.gsub(':', '')
        end

        @content = new_content if @content != new_content

        # Now find all references.
        new_references = new_content[1..-1].scan(CELL_REF_REG_EXP).inject Set.new do |memo, ref|
          memo << spreadsheet.find_or_create_cell(ref)
        end
      end

      add_references    new_references - old_references
      remove_references old_references - new_references

      eval true
    end

    def eval(reevaluate = false)
      previous_content = @evaluated_content

      @evaluated_content = nil if reevaluate

      @evaluated_content ||=
        if is_formula?
          puts ">>> Calculating formula for #{self.ref}" if DEBUG

          @last_evaluated_at = Time.now

          evaluated_content = content[1..-1]

          references.each do |cell|
            evaluated_content.gsub! cell.ref.to_s, cell.eval.to_s
          end

          Formula.instance_eval { eval evaluated_content }
        else
          content
        end

      # Fire all observers if evaluated content has changed.
      fire_observers if previous_content != @evaluated_content

      @evaluated_content || DEFAULT_VALUE
    end

    def directly_or_indirectly_references?(cell)
      cell == self || references.include?(cell) || references.any? { |reference| reference.directly_or_indirectly_references?(cell) }
    end

    def copy_to(dest_ref)
      dest_content = content.clone

      references.each do |reference|
        dest_content.gsub! reference.ref.to_s, reference.new_ref(ref, dest_ref).to_s
      end

      spreadsheet.set dest_ref, dest_content
    end

    def move_to(dest_ref)
      source_ref = ref
      @ref    = dest_ref

      spreadsheet.update_cell_ref source_ref, dest_ref

      observers.each do |observer|
        observer.update_reference source_ref, dest_ref
      end
    end

    def update_reference(old_ref, new_ref)
      self.content = self.content.gsub(old_ref.to_s, new_ref.to_s)
    end

    def new_ref(source_ref, dest_ref)
      ref_col, ref_row       = self.class.get_column_and_row(ref)
      source_col, source_row = self.class.get_column_and_row(source_ref)
      dest_col, dest_row     = self.class.get_column_and_row(dest_ref)

      col_diff = COL_RANGE.index(dest_col) - COL_RANGE.index(source_col)
      row_diff = dest_row - source_row

      new_col = COL_RANGE[COL_RANGE.index(ref_col) + col_diff]
      new_row = ref_row + row_diff

      "#{new_col}#{new_row}".to_sym
    end

    private

    def fire_observers
      puts "Firing #{ref}'s observers" if DEBUG && observers.any?

      observers.each do |cell|
        cell.eval true
      end
    end

    def add_reference(cell)
      if cell.directly_or_indirectly_references?(self)
        raise CircularReferenceError, "Circular reference detected when adding reference #{cell.ref} to #{ref}!"
      end

      puts "Adding reference #{cell.ref} to #{ref}" if DEBUG

      references << cell
      cell.add_observer self
    end

    def remove_reference(cell)
      puts "Removing reference #{cell.ref} from #{ref}" if DEBUG

      references.delete cell
      cell.remove_observer self
    end

    def add_references(cells)
      cells.each do |cell|
        add_reference cell
      end
    end

    def remove_references(cells)
      cells.each do |cell|
        remove_reference cell
      end
    end

    def is_formula?
      String === content && content[0] == '='
    end

    def self.splat_range(upper_left_ref, lower_right_ref)
      ul_col, ul_row = get_column_and_row(upper_left_ref)
      lr_col, lr_row = get_column_and_row(lower_right_ref)

      (ul_row..lr_row).map do |row|
        (ul_col..lr_col).map do |col|
          "#{col}#{row}".to_sym
        end
      end
    end

    def self.get_column_and_row(ref)
      ref.to_s =~ CELL_REF2_REG_EXP && [$1.upcase.to_sym, $2.to_i]
    end
  end

  class Formula
    def self.sum(*cell_values)
      puts "Calling sum() for #{cell_values.inspect}" if DEBUG

      cell_values.flatten.inject :+
    end
  end
end
