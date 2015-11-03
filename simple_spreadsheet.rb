DEBUG = false

require 'set'

class Class
  def delegate(method, options)
    define_method method do |*args|
      if (receiver = send(options[:to]))
        receiver.send method, *args
      end
    end
  end
end

class CellRef
  COL_RANGE          = ('A'..'ZZZ').to_a.map(&:to_sym)
  CELL_REF1          = '\b[A-Z]+[1-9]\d*\b'
  CELL_REF2          = '\b([A-Z]+)([1-9]\d*)\b'
  CELL_REF_REG_EXP   = /#{CELL_REF1}/i
  CELL_REF2_REG_EXP  = /#{CELL_REF2}/i
  CELL_RANGE_REG_EXP = /((#{CELL_REF1}):(#{CELL_REF1}))/i

  # List of possible exceptions.
  class IllegalCellReference < StandardError; end

  attr_reader :ref

  delegate :col_ref_index, to: :class
  delegate :col_ref_name,  to: :class

  def initialize(*ref)
    ref = [*ref].flatten

    ref[0] = col_ref_name(ref[0]) if Fixnum === ref[0]

    ref = ref.join
    ref = normalize_ref(ref)

    raise IllegalCellReference unless ref =~ /^#{CellRef::CELL_REF1}$/i

    @ref = ref
  end

  def col
    @col ||= col_and_row[0]
  end

  def col_index
    @col_index ||= col_ref_index col
  end

  def row
    @row ||= col_and_row[1]
  end

  def col_and_row
    @col_and_row ||= ref.to_s =~ CELL_REF2_REG_EXP && [$1.upcase.to_sym, $2.to_i]
  end

  def neighbor(col_count: 0, row_count: 0)
    raise IllegalCellReference unless col_index + col_count > 0 && row + row_count > 0

    CellRef.new col_ref_name(col_index + col_count), row + row_count
  end

  def right_neighbor(count = 1)
    neighbor col_count: count
  end

  def left_neighbor(count = 1)
    neighbor col_count: -count
  end

  def upper_neighbor(count = 1)
    neighbor row_count: -count
  end

  def lower_neighbor(count = 1)
    neighbor row_count: count
  end

  def to_s
    ref.to_s
  end

  def ==(other_ref)
    ref == (CellRef === other_ref ? other_ref.ref : normalize_ref(other_ref))
  end

  private

  def normalize_ref(ref)
    ref.upcase.to_sym
  end

  def self.col_ref_index(col_ref)
    COL_RANGE.index(col_ref.to_sym) + 1
  end

  def self.col_ref_name(col_index)
    COL_RANGE[col_index - 1]
  end

  def self.splat_range(upper_left_ref, lower_right_ref)
    upper_left_ref  = new(upper_left_ref)  unless self === upper_left_ref
    lower_right_ref = new(lower_right_ref) unless self === lower_right_ref

    ul_col, ul_row = upper_left_ref.col_and_row
    lr_col, lr_row = lower_right_ref.col_and_row

    (ul_row..lr_row).map do |row|
      (ul_col..lr_col).map do |col|
        new col, row
      end
    end
  end
end

class Cell
  DEFAULT_VALUE = 0

  # List of possible exceptions.
  class CircularReferenceError < StandardError; end

  attr_reader :spreadsheet, :ref, :references, :observers, :content, :raw_content, :last_evaluated_at

  def initialize(spreadsheet, ref, content = nil)
    puts "Creating cell #{ref}" if DEBUG

    ref = CellRef.new(ref) unless CellRef === ref

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
    puts "Setting new content `#{new_content}` to cell #{ref}" if DEBUG

    new_content = new_content.strip if String === new_content

    old_references = Set.new
    new_references = Set.new

    old_references = references.clone if has_formula?

    @raw_content, @content =
      if String === new_content
        [new_content.clone, new_content.clone]
      else
        [new_content, new_content]
      end

    if has_formula?
      # Splat ranges, e.g., replace 'A1:A3' by '[[A1, A2, A3]]'.
      @content[1..-1].scan(CellRef::CELL_RANGE_REG_EXP).each do |(range, upper_left_ref, lower_right_ref)|
        @content.gsub! /\b#{range}\b/i, CellRef.splat_range(upper_left_ref, lower_right_ref).flatten.map(&:to_s).to_s.gsub('"', '')
      end

      # Now find all references.
      new_references = find_references
    end

    add_references    new_references - old_references
    remove_references old_references - new_references

    eval true
  end

  def find_references
    if has_formula?
      content[1..-1].scan(CellRef::CELL_REF_REG_EXP).inject(Set.new) do |memo, ref|
        memo << spreadsheet.find_or_create_cell(ref)
      end
    else
      Set.new
    end
  end

  def eval(reevaluate = false)
    previous_content = @evaluated_content

    @evaluated_content = nil if reevaluate

    @evaluated_content ||=
      if has_formula?
        puts ">>> Calculating formula for #{self.ref}" if DEBUG

        @last_evaluated_at = Time.now

        evaluated_content = content[1..-1]

        references.each do |cell|
          evaluated_content.gsub! /\b#{cell.ref}\b/i, cell.eval.to_s
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
    cell == self ||
      references.include?(cell) ||
      references.any? { |reference| reference.directly_or_indirectly_references?(cell) }
  end

  def copy_to_range(dest_range)
    CellRef.splat_range(*dest_range.split(':')).flatten.each do |ref|
      copy_to ref
    end
  end

  def copy_to(dest_ref)
    dest_ref = CellRef.new(dest_ref) unless CellRef === dest_ref

    return if dest_ref === ref

    dest_content = raw_content.clone

    references.each do |reference|
      dest_content.gsub! /\b#{reference.ref.ref}\b/i, reference.new_ref(ref, dest_ref).ref.to_s
    end

    spreadsheet.set dest_ref, dest_content
  end

  def move_to!(dest_ref)
    dest_ref = CellRef.new(dest_ref) unless CellRef === dest_ref

    return if dest_ref === ref

    source_ref = ref
    @ref       = dest_ref

    spreadsheet.move_cell source_ref, dest_ref

    observers.each do |observer|
      observer.update_reference source_ref, dest_ref
    end
  end

  def move_right!(col_count = 1)
    move_to! ref.right_neighbor(col_count)
  end

  def move_left!(col_count = 1)
    move_to! ref.left_neighbor(col_count)
  end

  def move_down!(row_count = 1)
    move_to! ref.lower_neighbor(row_count)
  end

  def move_up!(row_count = 1)
    move_to! ref.upper_neighbor(row_count)
  end

  def update_reference(old_ref, new_ref)
    self.content = self.content.gsub(/\b#{old_ref}\b/i, new_ref.to_s)
  end

  # Calculates a cell's new reference when an observer cell is copied from `observer_source_ref` to `observer_dest_ref`.
  def new_ref(observer_source_ref, observer_dest_ref)
    col_diff = observer_dest_ref.col_index  - observer_source_ref.col_index
    row_diff = observer_dest_ref.row        - observer_source_ref.row

    ref.right_neighbor(col_diff).lower_neighbor(row_diff)
  end

  def has_formula?
    String === content && content[0] == '='
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
end

class Formula
  def self.sum(*cell_values)
    puts "Calling sum() for #{cell_values.inspect}" if DEBUG

    cell_values.flatten.inject :+
  end
end

class Spreadsheet
  PP_CELL_SIZE     = 20
  PP_ROW_REF_SIZE  = 5
  PP_COL_DELIMITER = ' | '

  # List of possible exceptions.
  class AlreadyExistentCellError < StandardError; end

  attr_reader :cells

  def initialize
    @cells = {
      all: {},
      by_col: {},
      by_row: {}
    }
  end

  def cell_count
    cells[:all].values.compact.size
  end

  def find_or_create_cell(ref, content = nil)
    (find_cell_ref(ref) || add_cell(ref)).tap do |cell|
      cell.content = content if content
    end
  end

  alias_method :set, :find_or_create_cell

  def add_cell(ref, content = nil)
    raise AlreadyExistentCellError, "Cell #{ref} already exists" if find_cell_ref(ref)

    Cell.new(self, ref, content).tap do |cell|
      update_cell_ref ref, cell
    end
  end

  def move_cell(old_ref, new_ref)
    cell = find_cell_ref(old_ref)

    update_cell_ref old_ref, nil
    update_cell_ref new_ref, cell
  end

  def add_col(col)
    cells[:by_col].find_all { |(col_ref, _)|
      CellRef.col_ref_index(col_ref) >= CellRef.col_ref_index(col)
    }.map { |(col_ref, col_cells)|
      { CellRef.col_ref_index(col_ref) => col_cells }
    }.sort { |a, b|
      b.keys[0] <=> a.keys[0]     # Descending order.
    }.each do |item|
      col_cells = item.values[0]
      col_cells.each &:move_right!
    end
  end

  def add_row(row)
    lower_rows = cells[:by_row].find_all do |(row_ref, _)|
      row_ref >= row
    end

    lower_rows.sort.each do |(_, row_cells)|
      row_cells.each &:move_down!
    end
  end

  def consistent?
    cells[:all].all? do |(_, cell)|
      consistent =
        if cell.has_formula?
          cell.find_references == cell.references && cell.references.all? do |reference|
            reference.observers.include? cell
          end
        else
          cell.references.empty?
        end

      next false unless consistent

      consistent = cell.observers.all? do |observer|
        observer.references.include? cell
      end

      next false unless consistent

      col, row = cell.ref.col_and_row

      consistent =
        cells[:by_col][col] && cells[:by_col][col][row] == cell &&
        cells[:by_row][row] && cells[:by_row][row][col] == cell

      next false unless consistent

      true
    end
  end

  def pp
    sorted_cols = cells[:by_col].map { |k, v| { CellRef.col_ref_index(k) => v } }.sort { |a, b| a.keys[0] <=> b.keys[0] }
    sorted_rows = cells[:by_row].sort

    max_col, _ = (max = sorted_cols.max { |a, b| a.keys[0] <=> b.keys[0] }) && max.keys[0]
    max_row, _ = sorted_rows.max

    if max_col && max_row
      print ' ' * PP_ROW_REF_SIZE
      puts (1..max_col).map { |col| CellRef.col_ref_name(col).to_s.rjust(PP_CELL_SIZE) }.join(PP_COL_DELIMITER)

      print ' ' * PP_ROW_REF_SIZE
      max_col.times do |i|
        print '-' * (PP_CELL_SIZE + 1 + (i == 0 ? 0 : 1))
        print '+' if i < max_col - 1
      end
      puts

      (1..max_row).each do |row|
        print "#{row}:".rjust(PP_ROW_REF_SIZE)

        (1..max_col).each  do |col|
          print PP_COL_DELIMITER if col > 1

          if (cell = cells[:by_row][row] && cells[:by_row][row][CellRef.col_ref_name(col)])
            print ((cell.has_formula? ? "[`#{cell.raw_content}`] " : '') + cell.eval.to_s).rjust(PP_CELL_SIZE)
          else
            print ' ' * PP_CELL_SIZE
          end
        end

        puts
      end
    else
      puts 'Empty spreadsheet'
    end

    nil
  end

  def repl
    read_value = -> (message, constraint = nil, default_value = nil) do
      default_value = default_value.to_s if default_value

      value = nil

      loop do
        print message
        value = gets.chomp
        value = nil if value == ''

        value ||= default_value

        valid_value =
          if !constraint
            true
          elsif value
            if value =~ /^\d+$/ && constraint.respond_to?(:include?)
              constraint.include? value.to_i
            elsif Regexp === constraint
              constraint =~ value
            elsif constraint.respond_to?(:include?)
              constraint.include? value.upcase
            end
          end

        break if valid_value
      end

      value
    end

    read_cell_ref = -> (message = 'Enter cell reference: ') do
      ref = read_value.call(message, /^#{CellRef::CELL_REF1}$/i)
    end

    read_cell_range = -> (message = 'Enter cell range: ') do
      ref = read_value.call(message, CellRef::CELL_RANGE_REG_EXP)
    end

    read_number = -> (message, default_value = nil) do
      number = read_value.call(message, 1..2**32, default_value)
      number.to_i
    end

    loop do
      begin
        ref = nil

        action = read_value.call(
          "Enter action [S - Set cell (default); M - Move cell; CC - Copy cell to cell; CR - Copy cell to range; AR - Add row; AC - Add col; Q - Quit]: ",
          ['S', 'M', 'CC', 'CR', 'AR', 'AC', 'Q'],
          'S'
        )

        case action.upcase
          when 'S' then
            ref     = read_cell_ref.call
            content = read_value.call("Enter content (for formulas start with a '='): ")

            set ref, content

          when 'M' then
            subaction = read_value.call(
              'Enter sub action [S - Specific position (default); U - Up; D - Down; L - Left; R - Right]: ',
              ['S', 'U', 'D', 'L', 'R'],
              'S'
            )

            ref = read_cell_ref.call('Select source reference: ')

            cell = find_or_create_cell(ref)

            case subaction.upcase
              when 'S' then
                cell.move_to! read_cell_ref.call('Select destination reference: ')
              when 'U' then
                cell.move_up! read_number.call('Enter # of rows (default: 1): ', 1)
              when 'D' then
                cell.move_down! read_number.call('Enter # of rows (default: 1): ', 1)
              when 'L' then
                cell.move_left! read_number.call('Enter # of cols (default: 1): ', 1)
              when 'R' then
                cell.move_right! read_number.call('Enter # of cols (default: 1): ', 1)
            end

          when 'CC' then
            ref      = read_cell_ref.call('Select source reference: ')
            cell     = find_or_create_cell(ref)
            dest_ref = read_cell_ref.call('Select destination reference: ')

            cell.copy_to dest_ref

          when 'CR' then
            ref        = read_cell_ref.call('Select source reference: ')
            cell       = find_or_create_cell(ref)
            dest_range = read_cell_range.call('Select destination range: ')

            cell.copy_to_range dest_range

          when 'AR' then
            add_row read_number.call('Enter row number (>= 1): ')

          when 'AC' then
            col = read_value.call('Enter col name (>= "A"): ', 'A'..'ZZZ')

            add_col CellRef.col_ref_index(col)

          when 'Q' then
            break;

          else
            next;
        end
      rescue StandardError => e
        puts "Error: `#{e}`."
        puts 'Stack trace:'
        puts e.backtrace
      end

      pp
    end
  end

  private

  def find_cell_ref(ref)
    ref = CellRef.new(ref) unless CellRef === ref

    cells[:all][ref.ref]
  end

  def update_cell_ref(ref, cell)
    ref = CellRef.new(ref) unless CellRef === ref

    col, row = ref.col_and_row

    cells[:by_col][col] ||= {}
    cells[:by_row][row] ||= {}

    cells[:all][ref.ref] = cells[:by_col][col][row] = cells[:by_row][row][col] = cell
  end
end

def run!
  Spreadsheet.new.repl
end

run! if __FILE__ == $0
