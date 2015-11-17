DEBUG = false

require 'colorize'

class Class
  def delegate(*args)
    options = args.last.is_a?(Hash) ? args.pop : {}
    methods = args

    raise ArgumentError, ':to option is mandatory' unless options[:to]

    methods.each do |method|
      define_method method do |*args, &block|
        if (receiver = send(options[:to]))
          receiver.send method, *args, &block
        end
      end
    end
  end

  def delegate_all(options = {})
    raise ArgumentError, ':to option is mandatory' unless options[:to]

    define_method :method_missing do |method, *args, &block|
      if (receiver = send(options[:to]))
        if receiver.respond_to?(method)
          receiver.send method, *args, &block
        else
          super method, *args, &block
        end
      end
    end
  end
end

class Array
  def subtract(another_array)
    reject { |item| another_array.include?(item) }
  end

  def unique_add(item)
    self << item unless include?(item)
  end
end

class String
  def truncate(limit, delimiter = '...')
    return self if limit < 0

    if size <= limit
      self
    else
      truncate_position = limit - 1 - delimiter.size

      if truncate_position >= 0
        self[0..truncate_position] + delimiter
      else
        delimiter
      end
    end
  end
end

class CellRef
  COL_RANGE                      = ('A'..'ZZZ').to_a.map(&:to_sym)
  CELL_REF_FOR_RANGES            = '[A-Z]+[1-9]\d*'
  CELL_REF                       = '\$?[A-Z]+\$?[1-9]\d*'
  CELL_REF_WITH_PARENS           = '(\$?)([A-Z]+)(\$?)([1-9]\d*)'
  CELL_REF_REG_EXP               = /#{CELL_REF}/i
  CELL_REF_WITH_PARENS_REG_EXP   = /#{CELL_REF_WITH_PARENS}/i
  CELL_RANGE                     = "#{CELL_REF_FOR_RANGES}:#{CELL_REF_FOR_RANGES}"
  CELL_RANGE_WITH_PARENS         = "(#{CELL_REF_FOR_RANGES}):(#{CELL_REF_FOR_RANGES})"
  CELL_RANGE_WITH_PARENS_REG_EXP = /(#{CELL_RANGE_WITH_PARENS})/i

  # List of possible exceptions.
  class IllegalCellReference < StandardError; end

  attr_reader :ref

  delegate :col_ref_index, :col_ref_name, :normalize_ref, :parse_ref, to: :class

  def initialize(*ref)
    ref.flatten!
    ref[0] = col_ref_name(ref[0]) if ref[0].is_a?(Fixnum)
    ref    = ref.join

    raise IllegalCellReference unless ref =~ /^#{CellRef::CELL_REF}$/i

    @ref = normalize_ref(ref)
  end

  def self.self_or_new(ref_or_cell_ref)
    if ref_or_cell_ref.is_a?(CellRef)
      ref_or_cell_ref
    else
      new ref_or_cell_ref
    end
  end

  def col
    @col ||= col_and_row[0]
  end

  def col_index
    @col_index ||= col_ref_index(col)
  end

  def row
    @row ||= col_and_row[1]
  end

  def col_and_row
    @col_and_row ||= parse_ref(ref)
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

  def to_sym
    ref.to_sym
  end

  def to_s
    ref.to_s
  end

  def ==(other_ref)
    ref == (other_ref.is_a?(CellRef) ? other_ref.ref : normalize_ref(other_ref))
  end

  def self.normalize_ref(ref)
    parse_ref(ref).join.to_sym
  end

  def self.parse_ref(ref)
    ref.to_s =~ CELL_REF_WITH_PARENS_REG_EXP && [$2.upcase.to_sym, $4.to_i]
  end

  def self.col_ref_index(col_ref)
    COL_RANGE.index(col_ref.upcase.to_sym) + 1
  end

  def self.col_ref_name(col_index)
    COL_RANGE[col_index - 1]
  end

  def self.splat_range(upper_left_ref, lower_right_ref)
    upper_left_ref  = new(upper_left_ref)  unless upper_left_ref.is_a?(self)
    lower_right_ref = new(lower_right_ref) unless lower_right_ref.is_a?(self)

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

    ref = CellRef.self_or_new(ref)

    @spreadsheet = spreadsheet
    @ref         = ref
    @references  = []
    @observers   = []

    self.content = content
  end

  def add_observer(observer)
    puts "Adding observer #{observer.ref} to #{ref}" if DEBUG

    observers.unique_add observer
  end

  def remove_observer(observer)
    puts "Removing observer #{observer.ref} from #{ref}" if DEBUG

    observers.delete observer
  end

  def content=(new_content)
    puts "Replacing content `#{content}` with new content `#{new_content}` in cell #{ref}" if DEBUG

    new_content = new_content.strip if new_content.is_a?(String)

    old_references = references.clone
    new_references = []

    if new_content.is_a?(String)
      uppercased_content = new_content.gsub(/(#{CellRef::CELL_REF})/i) { $1.upcase }

      @raw_content, @content = uppercased_content, uppercased_content.clone
    else
      @raw_content, @content = new_content, new_content
    end

    if formula?
      # Splat ranges, e.g., replace 'A1:A3' by '[[A1, A2, A3]]'.
      @content[1..-1].scan(CellRef::CELL_RANGE_WITH_PARENS_REG_EXP).each do |(range, upper_left_ref, lower_right_ref)|
        @content.gsub!  /(?<![A-Z])#{range}(?![1-9])/i,
                        CellRef.splat_range(upper_left_ref, lower_right_ref).flatten.map(&:to_s).to_s.gsub('"', '')
      end

      new_references = find_references
    end

    begin
      add_references    new_references.subtract(old_references)   # Do not use `new_references - old_references`
      remove_references old_references.subtract(new_references)   # and `old_references - new_references`.

      eval true
    rescue StandardError => e
      @evaluated_content, @raw_content, @content = "Error '#{e.message}': `#{@content}`"

      remove_all_references
    end
  end

  def find_references
    if formula?
      content[1..-1].scan(CellRef::CELL_REF_REG_EXP).map { |ref|
        cell = spreadsheet.find_or_create_cell(ref)

        CellWrapper.new cell, ref
      }.uniq
    else
      []
    end
  end

  def eval(reevaluate = false)
    previous_content = @evaluated_content

    @evaluated_content = nil if reevaluate

    @evaluated_content ||=
      if formula?
        puts ">>> Calculating formula for #{self.ref}" if DEBUG

        @last_evaluated_at = Time.now

        evaluated_content = content[1..-1]

        references.each do |reference|
          # Replace the reference in the content, making sure it's not preceeded by a letter or succeeded by a number. This simple
          # rule assures references like 'A1' are correctly replaced in formulas like '= A1 + A11 * AA1 / AA11'
          evaluated_content.gsub! /(?<![A-Z\$])#{Regexp.escape(reference.full_ref)}(?![1-9])/i, reference.eval.to_s
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
    dest_ref = CellRef.self_or_new(dest_ref)

    return if dest_ref == ref

    dest_content = raw_content.clone

    references.each do |reference|
      dest_content.gsub! /(?<![A-Z])#{Regexp.escape(reference.full_ref)}(?![1-9])/i, reference.new_ref(ref, dest_ref)
    end

    spreadsheet.set dest_ref, dest_content
  end

  def move_to!(dest_ref)
    dest_ref = CellRef.self_or_new(dest_ref)

    return if dest_ref == ref

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
    puts "Replacing reference `#{old_ref}` with `#{new_ref}` in #{ref}" if DEBUG

    old_col, old_row = CellRef.parse_ref(old_ref)
    new_col, new_row = CellRef.parse_ref(new_ref)

    self.content = self.content.gsub(/(?<![A-Z])(\$?)#{old_col}(\$?)#{old_row}(?![1-9])/i) { [$1, new_col, $2, new_row].join }
  end

  def formula?
    content.is_a?(String) && content[0] == '='
  end

  def ==(another_cell_or_cell_wrapper)
    if another_cell_or_cell_wrapper.is_a?(CellWrapper)
      self == another_cell_or_cell_wrapper.cell
    else
      super
    end
  end

  def inspect
    {
      spreadsheet: spreadsheet.object_id,
      ref:         ref.ref,
      references:  references,
      observers:   observers
    }
  end

  private

  def fire_observers
    puts "Firing #{ref}'s observers" if DEBUG && observers.any?

    observers.each do |observer|
      observer.eval true
    end
  end

  def add_reference(reference)
    if reference.directly_or_indirectly_references?(self)
      raise CircularReferenceError, "Circular reference detected when adding reference #{reference.ref} to #{ref}!"
    end

    puts "Adding reference #{reference.ref} to #{ref}" if DEBUG

    references.unique_add reference
    reference.add_observer self
  end

  def remove_all_references
    references.each do |reference|
      remove_reference reference
    end
  end

  def remove_reference(reference)
    puts "Removing reference #{reference.ref} from #{ref}" if DEBUG

    references.delete reference
    reference.remove_observer(self) unless references.any? { |ref| ref.cell == reference.cell }
  end

  def add_references(new_references)
    new_references.each do |reference|
      add_reference reference
    end
  end

  def remove_references(old_references)
    old_references.each do |reference|
      remove_reference reference
    end
  end
end

class CellWrapper
  attr_reader :cell

  # The code below could also re written as: `delegate_all to: :cell`, which would be a little slower (due to the use of the
  # :method_missing method, but more generic.
  delegate :directly_or_indirectly_references?, :ref, :eval, :observers, :add_observer, :remove_observer, :spreadsheet, to: :cell

  def initialize(cell, ref)
    raise IllegalCellReference unless ref.to_s =~ CellRef::CELL_REF_WITH_PARENS_REG_EXP

    @cell            = cell
    @is_absolute_col = $1 == '$'
    @is_absolute_row = $3 == '$'
  end

  def absolute_col?
    @is_absolute_col
  end

  def absolute_row?
    @is_absolute_row
  end

  def full_ref
    col, row = cell.ref.col_and_row

    ref_parts = []
    ref_parts << '$' if absolute_col?
    ref_parts << col
    ref_parts << '$' if absolute_row?
    ref_parts << row

    ref_parts.join
  end

  # Calculates a cell's new reference when an observer cell is copied from `observer_source_ref` to `observer_dest_ref`.
  def new_ref(observer_source_ref, observer_dest_ref)
    col_diff = absolute_col? ? 0 : observer_dest_ref.col_index  - observer_source_ref.col_index
    row_diff = absolute_row? ? 0 : observer_dest_ref.row        - observer_source_ref.row

    target_ref = ref.right_neighbor(col_diff).lower_neighbor(row_diff)

    ref_parts = []
    ref_parts << '$' if absolute_col?
    ref_parts << target_ref.col
    ref_parts << '$' if absolute_row?
    ref_parts << target_ref.row

    ref_parts.join
  end

  def ==(another_cell_or_cell_wrapper)
    # puts "#== called for #{another_cell_or_cell_wrapper}"

    case another_cell_or_cell_wrapper
    when CellWrapper then
      cell == another_cell_or_cell_wrapper.cell &&
        absolute_col? == another_cell_or_cell_wrapper.absolute_col? &&
        absolute_row? == another_cell_or_cell_wrapper.absolute_row?
    when Cell then
      cell == another_cell_or_cell_wrapper
    when String, Symbol then
      full_ref == another_cell_or_cell_wrapper.to_s.upcase
    else
      false
    end
  end

  def coerce(any_object)
    # puts "#coerce called for #{any_object}"

    case any_object
    when String, Symbol then
      # cell = spreadsheet.find_or_create_cell(any_object)
      #
      # [CellWrapper.new(cell, any_object), self]

      [any_object.to_s.upcase, full_ref]
    else
      super
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
  PP_CELL_SIZE     = 30
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

  def set(ref, content)
    find_or_create_cell ref, content
  end

  def add_cell(ref, content = nil)
    raise AlreadyExistentCellError, "Cell #{ref} already exists" if find_cell_ref(ref)

    Cell.new(self, ref, content).tap do |cell|
      update_cell_ref ref, cell
    end
  end

  def move_cell(old_ref, new_ref)
    cell = find_cell_ref(old_ref)

    delete_cell_ref old_ref
    update_cell_ref new_ref, cell
  end

  def add_col(col_to_add, count = 1)
    col_to_add = CellRef.col_ref_index(col_to_add) unless col_to_add.is_a?(Fixnum)

    cells[:by_col].select { |(col, _)| col >= col_to_add }.sort.reverse.each do |(_, rows)|
      rows.sort.each { |(_, cell)| cell.move_right! count }
    end
  end

  def delete_col(col_to_delete, count = 1)
    col_to_delete = CellRef.col_ref_index(col_to_delete) unless col_to_delete.is_a?(Fixnum)

    cells[:by_col][col_to_delete].each { |(_, cell)| delete_cell_ref cell.ref }

    cells[:by_col].select { |(col, _)| col >= col_to_delete + count }.sort.each do |(_, rows)|
      rows.sort.each { |(_, cell)| cell.move_left! count }
    end
  end

  def add_row(row_to_add, count = 1)
    cells[:by_row].select { |(row, _)| row >= row_to_add }.sort.reverse.each do |(_, cols)|
      cols.sort.each { |(_, cell)| cell.move_down! count }
    end
  end

  def delete_row(row_to_delete, count = 1)
    cells[:by_row][row_to_delete].each { |(_, cell)| delete_cell_ref cell.ref }

    cells[:by_row].select { |(row, _)| row >= row_to_delete + count }.sort.each do |(_, cols)|
      cols.sort.each { |(_, cell)| cell.move_up! count }
    end
  end

  def consistent?
    cells[:all].all? do |(_, cell)|
      consistent =
        if cell.formula?
          cell_references = cell.find_references.inject [] do |memo, reference|
            memo.unique_add reference
            memo
          end

          cell_references == cell.references && cell.references.all? do |reference|
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

      col = cell.ref.col_index
      row = cell.ref.row

      consistent =
        cells[:by_col][col] && cells[:by_col][col][row] == cell &&
        cells[:by_row][row] && cells[:by_row][row][col] == cell

      next false unless consistent

      true
    end
  end

  def pp(last_change)
    lrjust = -> (ltext, rtext, size) do
      spacing_size = size - (ltext.size + rtext.size)

      if spacing_size >= 0
        ltext + ' ' * spacing_size + rtext
      else
        lrjust.call ltext.truncate(ltext.size + spacing_size - 1), rtext, size
      end
    end

    max_col = (max = cells[:by_col].sort.max) && max[0]
    max_row = (max = cells[:by_row].sort.max) && max[0]

    if max_col && max_row
      print ' '
      print ' ' * PP_ROW_REF_SIZE
      puts (1..max_col).map { |col| CellRef.col_ref_name(col).to_s.rjust(PP_CELL_SIZE) }.join(PP_COL_DELIMITER)

      print ' '
      print ' ' * PP_ROW_REF_SIZE
      max_col.times do |i|
        print '-' * (PP_CELL_SIZE + 1 + (i == 0 ? 0 : 1))
        print '+' if i < max_col - 1
      end
      puts

      (1..max_row).each do |row|
        print "#{row}:".rjust(PP_ROW_REF_SIZE)
        print ' '

        (1..max_col).each  do |col|
          print PP_COL_DELIMITER if col > 1

          if (cell = cells[:by_row][row] && cells[:by_row][row][col])
            value = cell.eval

            highlight_cell = false

            text =
              if cell.formula?
                # Highlight cell if value has changed.
                highlight_cell = last_change && cell.last_evaluated_at > last_change

                lrjust.call("`#{cell.raw_content}`", value.to_s, PP_CELL_SIZE)
              else
                value.to_s.rjust(PP_CELL_SIZE)
              end

            print highlight_cell ? text.truncate(PP_CELL_SIZE).blue.on_light_white : text.truncate(PP_CELL_SIZE)
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
            elsif constraint.is_a?(Regexp)
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
      ref = read_value.call(message, /^#{CellRef::CELL_REF}$/i)
    end

    read_cell_range = -> (message = 'Enter cell range: ') do
      ref = read_value.call(message, /^#{CellRef::CELL_RANGE}$/i)
    end

    read_number = -> (message, default_value = nil) do
      number = read_value.call(message, 1..2**32, default_value)
      number.to_i
    end

    last_change = nil

    loop do
      begin
        ref = nil

        pp last_change

        last_change = Time.now

        action = read_value.call(
          "Enter action [S - Set cell (default); M - Move cell; CC - Copy cell to cell; CR - Copy cell to range; AC - Add col; AR - Add row; DC - Delete col; DR - Delete row; Q - Quit]: ",
          ['S', 'M', 'CC', 'CR', 'AR', 'AC', 'DC', 'DR', 'Q'],
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

        when 'AC' then
          col       = read_value.call('Enter col name (>= "A"): ', 'A'..'ZZZ')
          col_count = read_number.call('Enter # of cols (default: 1): ', 1)

          add_col col, col_count

        when 'AR' then
          row       = read_number.call('Enter row #: ')
          row_count = read_number.call('Enter # of rows (default: 1): ', 1)

          add_row row, row_count

        when 'DC' then
          col       = read_value.call('Enter col name (>= "A"): ', 'A'..'ZZZ')
          col_count = read_number.call('Enter # of cols (default: 1): ', 1)

          delete_col col, col_count

        when 'DR' then
          row       = read_number.call('Enter row #: ')
          row_count = read_number.call('Enter # of rows (default: 1): ', 1)

          delete_row row, row_count

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
    end
  end

  private

  def find_cell_ref(ref)
    ref = CellRef.self_or_new(ref)

    cells[:all][ref.ref]
  end

  def delete_cell_ref(ref)
    ref = CellRef.self_or_new(ref)

    col = ref.col_index
    row = ref.row

    cells[:by_col][col] ||= {}
    cells[:by_row][row] ||= {}

    cells[:all].delete ref.ref
    cells[:by_col][col].delete row
    cells[:by_row][row].delete col
  end

  def update_cell_ref(ref, cell)
    ref = CellRef.self_or_new(ref)

    col = ref.col_index
    row = ref.row

    cells[:by_col][col] ||= {}
    cells[:by_row][row] ||= {}

    cells[:all][ref.ref] = cells[:by_col][col][row] = cells[:by_row][row][col] = cell
  end
end

def run!
  spreadsheet = Spreadsheet.new

  # a1 = spreadsheet.set(:A1, 'BRL/Dollar rate:')
  # b1 = spreadsheet.set(:A2, 3.90)
  #
  # b3 = spreadsheet.set(:B3, 'Expenses (in USD)')
  # c3 = spreadsheet.set(:C3, 'Expenses (in BRL)')
  #
  # a4 = spreadsheet.set(:A4, 'Rent')
  # a5 = spreadsheet.set(:A5, 'Payroll')
  # a6 = spreadsheet.set(:A5, 'Utilities')
  #
  # b4 = spreadsheet.set(:B4, 10.00)
  # b5 = spreadsheet.set(:B5, 20.00)
  # b6 = spreadsheet.set(:B6, 30.00)
  #
  # c4 = spreadsheet.set(:C4, '= B4 * $A$2')
  #
  # c4.copy_to_range('C5:C6')

  # a1 = spreadsheet.set(:A1, 1)
  # a2 = spreadsheet.set(:A2, 2)
  # a3 = spreadsheet.set(:A3, 4)
  # a4 = spreadsheet.set(:A4, 8)
  # a5 = spreadsheet.set(:A5, 16)
  #
  # 6.upto(40) do |i|
  #   formula = (rand(i - 1) + 1).times.map do |j|
  #     operator = j == 0 ? '' : '+'    # ['+', '-', '*'].sample
  #
  #     operator + [
  #       rand(2) == 0 ? '' : '$',
  #       'A',
  #       rand(2) == 0 ? '' : '$',
  #       rand(i - 1) + 1
  #     ].join
  #   end
  #
  #   spreadsheet.set [:A, i], "= #{formula.join}"
  # end

  # Fibonacci sequence.
  a1 = spreadsheet.set(:A1, 1)
  a2 = spreadsheet.set(:A2, 1)
  a3 = spreadsheet.set(:A3, '= $A$1 + $A$2')

  4.upto(30) do |i|
    # spreadsheet.set [:A, i], "= A#{i - 1} + A#{i - 2}"
    spreadsheet.set [:A, i], "= sum(A#{i - 2}:A#{i - 1})"
  end

  spreadsheet.repl
end

run! if __FILE__ == $0
