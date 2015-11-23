require 'colorize'

class Object
  DEBUG = false

  def log(msg)
    puts "[#{Time.now}] #{msg}" if DEBUG
  end
end

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
    self
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

class CellCoordinate
  COL_RANGE                      = ('A'..'ZZZ').to_a.map(&:to_sym)
  CELL_COORD_FOR_RANGES          = '[A-Z]+[1-9]\d*'
  CELL_COORD                     = '\$?[A-Z]+\$?[1-9]\d*'
  CELL_COORD_WITH_PARENS         = '(\$?)([A-Z]+)(\$?)([1-9]\d*)'
  CELL_COORD_REG_EXP             = /#{CELL_COORD}/i
  CELL_COORD_WITH_PARENS_REG_EXP = /#{CELL_COORD_WITH_PARENS}/i
  CELL_RANGE                     = "#{CELL_COORD_FOR_RANGES}:#{CELL_COORD_FOR_RANGES}"
  CELL_RANGE_WITH_PARENS         = "(#{CELL_COORD_FOR_RANGES}):(#{CELL_COORD_FOR_RANGES})"
  CELL_RANGE_WITH_PARENS_REG_EXP = /(#{CELL_RANGE_WITH_PARENS})/i

  # List of possible exceptions.
  class IllegalCellReference < StandardError; end

  attr_reader :coord

  delegate :col_coord_index, :col_coord_name, :normalize_coord, :parse_coord, to: :class

  def initialize(*coord)
    coord.flatten!
    coord[0] = col_coord_name(coord[0]) if coord[0].is_a?(Fixnum)
    coord    = coord.join

    raise IllegalCellReference unless coord =~ /^#{CellCoordinate::CELL_COORD}$/i

    @coord = normalize_coord(coord)
  end

  def self.self_or_new(coord_or_cell_coord)
    if coord_or_cell_coord.is_a?(CellCoordinate)
      coord_or_cell_coord
    else
      new coord_or_cell_coord
    end
  end

  def col
    @col ||= col_and_row[0]
  end

  def col_index
    @col_index ||= col_coord_index(col)
  end

  def row
    @row ||= col_and_row[1]
  end

  def col_and_row
    @col_and_row ||= parse_coord(coord)
  end

  def neighbor(col_count: 0, row_count: 0)
    raise IllegalCellReference unless col_index + col_count > 0 && row + row_count > 0

    CellCoordinate.new col_coord_name(col_index + col_count), row + row_count
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
    coord.to_sym
  end

  def to_s
    coord.to_s
  end

  def ==(other_coord)
    coord == (other_coord.is_a?(CellCoordinate) ? other_coord.coord : normalize_coord(other_coord))
  end

  def self.normalize_coord(coord)
    parse_coord(coord).join.to_sym
  end

  def self.parse_coord(coord)
    coord.to_s =~ CELL_COORD_WITH_PARENS_REG_EXP && [$2.upcase.to_sym, $4.to_i]
  end

  def self.col_coord_index(col_coord)
    COL_RANGE.index(col_coord.upcase.to_sym) + 1
  end

  def self.col_coord_name(col_index)
    COL_RANGE[col_index - 1]
  end

  def self.splat_range(upper_left_coord, lower_right_coord)
    upper_left_coord  = new(upper_left_coord)  unless upper_left_coord.is_a?(self)
    lower_right_coord = new(lower_right_coord) unless lower_right_coord.is_a?(self)

    ul_col, ul_row = upper_left_coord.col_and_row
    lr_col, lr_row = lower_right_coord.col_and_row

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

  attr_reader :spreadsheet, :coord, :references, :observers, :content, :raw_content, :last_evaluated_at

  def initialize(spreadsheet, coord, content = nil)
    log "Creating cell #{coord}"

    coord = CellCoordinate.self_or_new(coord)

    @spreadsheet = spreadsheet
    @coord       = coord
    @references  = []
    @observers   = []

    self.content = content
  end

  def add_observer(observer)
    log "Adding observer #{observer.coord} to #{coord}"

    observers.unique_add observer
  end

  def remove_observer(observer)
    log "Removing observer #{observer.coord} from #{coord}"

    observers.delete observer
  end

  def content=(new_content)
    log "Replacing content `#{content}` with new content `#{new_content}` in cell #{coord}"

    new_content.strip! if new_content.is_a?(String)

    return if content == new_content

    old_references = references.clone
    new_references = []

    if new_content.is_a?(String)
      uppercased_content = new_content.gsub(/(#{CellCoordinate::CELL_COORD})/i) { $1.upcase }

      @raw_content, @content = uppercased_content, uppercased_content.clone
    else
      @raw_content, @content = new_content, new_content
    end

    @last_evaluated_at = nil    # Assume it was never evaluated before, since its contents have effectively changed.

    if formula?
      # Splat ranges, e.g., replace 'A1:A3' by '[[A1, A2, A3]]'.
      @content[1..-1].scan(CellCoordinate::CELL_RANGE_WITH_PARENS_REG_EXP).each do |(range, upper_left_coord, lower_right_coord)|
        @content.gsub!  /(?<![A-Z])#{range}(?![1-9])/i,
                        '[' + CellCoordinate.splat_range(upper_left_coord, lower_right_coord).map { |row|
                          '[' + row.map(&:to_s).join(', ') + ']'
                        }.join(', ') + ']'
      end

      new_references = find_references

      log "References found: #{new_references.map(&:full_coord).join(', ')}"
    end

    begin
      references_to_add    = new_references.subtract(old_references)   # Do not use `new_references - old_references`
      references_to_remove = old_references.subtract(new_references)   # and `old_references - new_references`.

      if references_to_add.any? || references_to_remove.any?
        # Notify all direct and indirect observers to reset their circular reference check cache (memoization). Notice
        # that this step must be done necessarily BEFORE adding or removing references, given observers
        reset_circular_reference_check_cache

        add_references    references_to_add
        remove_references references_to_remove
      end

      eval true
    rescue StandardError => e
      @evaluated_content, @raw_content, @content = "Error '#{e.message}': `#{@content}`"

      remove_all_references
    end
  end

  def find_references
    return [] unless formula?

    content[1..-1].scan(CellCoordinate::CELL_COORD_REG_EXP).inject [] do |memo, coord|
      cell           = spreadsheet.find_or_create_cell(coord)
      cell_reference = CellReference.new cell, coord

      memo.unique_add cell_reference
    end
  end

  def eval(reevaluate = false)
    previous_content = @evaluated_content

    if reevaluate
      if formula?
        latest_evaluated_reference_timestamp = references.map(&:last_evaluated_at).compact.max

        if latest_evaluated_reference_timestamp && last_evaluated_at && last_evaluated_at >= latest_evaluated_reference_timestamp
          log "Skipping reevaluation for #{coord}"
        else
          @evaluated_content = nil
        end
      else
        @evaluated_content = nil
      end
    end

    @evaluated_content ||= begin
      new_evaluated_content =
        if formula?
          log ">>> Calculating formula for #{coord}"

          evaluated_content = content[1..-1]

          references.each do |reference|
            # Replace the reference in the content, making sure it's not preceeded by a letter or succeeded by a number. This simple
            # rule assures references like 'A1' are correctly replaced in formulas like '= A1 + A11 * AA1 / AA11'
            evaluated_content.gsub! /(?<![A-Z\$])#{Regexp.escape(reference.full_coord)}(?![0-9])/i, reference.eval.to_s
          end

          # Evaluate the cell's content in the Formula context, so "functions" like `sum`, `average` etc are simply treated as calls to
          # Formula's (singleton) methods.
          Formula.instance_eval { eval evaluated_content }
        else
          content
        end

      @last_evaluated_at = Time.now

      new_evaluated_content
    end

    # Fire all observers if evaluated content has changed.
    fire_observers if previous_content != @evaluated_content

    @evaluated_content || DEFAULT_VALUE
  end

  def reset_circular_reference_check_cache
    return unless @directly_or_indirectly_references

    log "Resetting circular reference check cache for #{coord}"

    @directly_or_indirectly_references = nil

    reset_observers_circular_reference_check_cache
  end

  def reset_observers_circular_reference_check_cache
    observers.each &:reset_circular_reference_check_cache
  end

  def directly_or_indirectly_references?(cell)
    log "Checking if #{cell.coord} directly or indirectly references #{coord}"

    @directly_or_indirectly_references ||= {}

    if @directly_or_indirectly_references.has_key?(cell)
      @directly_or_indirectly_references[cell]
    else
      @directly_or_indirectly_references[cell] =
        cell == self ||
          references.include?(cell) ||
          references.any? { |reference| reference.directly_or_indirectly_references?(cell) }
    end
  end

  def copy_to_range(dest_range)
    CellCoordinate.splat_range(*dest_range.split(':')).flatten.each do |coord|
      copy_to coord
    end
  end

  def copy_to(dest_coord)
    dest_coord = CellCoordinate.self_or_new(dest_coord)

    return if dest_coord == coord

    if raw_content.is_a?(String)
      dest_content = raw_content.clone

      log "Content before replacements in copy_to: #{dest_content}"

      dest_content.gsub! CellCoordinate::CELL_COORD_WITH_PARENS_REG_EXP do |content_coord|
        if (cell = references.find { |reference| reference == content_coord })
          cell.new_coord coord, dest_coord
        end
      end

      log "Content after replacements in copy_to: #{dest_content}"
    else
      dest_content = raw_content
    end

    spreadsheet.set dest_coord, dest_content
  end

  def move_to!(dest_coord)
    dest_coord = CellCoordinate.self_or_new(dest_coord)

    return if dest_coord == coord

    source_coord = coord
    @coord       = dest_coord

    spreadsheet.move_cell source_coord, dest_coord

    observers.each do |observer|
      observer.update_reference source_coord, dest_coord
    end
  end

  def move_right!(col_count = 1)
    move_to! coord.right_neighbor(col_count)
  end

  def move_left!(col_count = 1)
    move_to! coord.left_neighbor(col_count)
  end

  def move_down!(row_count = 1)
    move_to! coord.lower_neighbor(row_count)
  end

  def move_up!(row_count = 1)
    move_to! coord.upper_neighbor(row_count)
  end

  def update_reference(old_coord, new_coord)
    log "Replacing reference `#{old_coord}` with `#{new_coord}` in #{coord}"

    old_col, old_row = CellCoordinate.parse_coord(old_coord)
    new_col, new_row = CellCoordinate.parse_coord(new_coord)

    # Do not use gsub! (since the setter won't be called).
    self.content = self.content.gsub(/(?<![A-Z])(\$?)#{old_col}(\$?)#{old_row}(?![1-9])/i) { [$1, new_col, $2, new_row].join }
  end

  def formula?
    content.is_a?(String) && content[0] == '='
  end

  def ==(another_cell_or_cell_reference)
    if another_cell_or_cell_reference.is_a?(CellReference)
      self == another_cell_or_cell_reference.cell
    else
      super
    end
  end

  def inspect
    {
      spreadsheet: spreadsheet.object_id,
      coord:       coord.coord,
      references:  references,
      observers:   observers
    }
  end

  private

  def fire_observers
    log "Firing #{coord}'s observers" if observers.any?

    observers.each do |observer|
      observer.eval true
    end
  end

  def add_reference(reference)
    if reference.directly_or_indirectly_references?(self)
      raise CircularReferenceError, "Circular reference detected when adding reference #{reference.coord} to #{coord}"
    end

    log "Adding reference #{reference.coord} to #{coord}"

    references.unique_add reference
    reference.add_observer self
  end

  def remove_all_references
    references.each do |reference|
      remove_reference reference
    end
  end

  def remove_reference(reference)
    log "Removing reference #{reference.coord} from #{coord}"

    references.delete reference
    reference.remove_observer(self) unless references.any? { |coord| coord.cell == reference.cell }
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

class CellReference
  attr_reader :cell

  # The code below could also re written as: `delegate_all to: :cell`, which would be more generic but a little slower (due to the use of
  # the :method_missing method).
  delegate :directly_or_indirectly_references?, :coord, :eval, :observers, :add_observer, :remove_observer, :spreadsheet, :last_evaluated_at, to: :cell

  def initialize(cell, coord)
    raise IllegalCellReference unless coord.to_s =~ CellCoordinate::CELL_COORD_WITH_PARENS_REG_EXP

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

  def full_coord
    col, row = cell.coord.col_and_row

    coord_parts = []
    coord_parts << '$' if absolute_col?
    coord_parts << col
    coord_parts << '$' if absolute_row?
    coord_parts << row

    coord_parts.join
  end

  # Calculates a cell's new reference when an observer cell is copied from `observer_source_coord` to `observer_dest_coord`.
  def new_coord(observer_source_coord, observer_dest_coord)
    col_diff = absolute_col? ? 0 : observer_dest_coord.col_index  - observer_source_coord.col_index
    row_diff = absolute_row? ? 0 : observer_dest_coord.row        - observer_source_coord.row

    target_coord = coord.right_neighbor(col_diff).lower_neighbor(row_diff)

    coord_parts = []
    coord_parts << '$' if absolute_col?
    coord_parts << target_coord.col
    coord_parts << '$' if absolute_row?
    coord_parts << target_coord.row

    coord_parts.join
  end

  def ==(another_cell_or_cell_reference)
    case another_cell_or_cell_reference
    when CellReference then
      cell == another_cell_or_cell_reference.cell &&
        absolute_col? == another_cell_or_cell_reference.absolute_col? &&
        absolute_row? == another_cell_or_cell_reference.absolute_row?
    when Cell then
      cell == another_cell_or_cell_reference
    when String, Symbol then
      full_coord == another_cell_or_cell_reference.to_s.upcase
    else
      false
    end
  end

  # def coerce(any_object)
  #   # log "#coerce called for #{any_object}"
  #
  #   case any_object
  #   when String, Symbol then
  #     # cell = spreadsheet.find_or_create_cell(any_object)
  #     #
  #     # [CellReference.new(cell, any_object), self]
  #
  #     [any_object.to_s.upcase, full_coord]
  #   else
  #     super
  #   end
  # end
end

class Formula
  def self.sum(*cell_values)
    log "Calling sum() for #{cell_values.inspect}"

    cell_values.flatten.inject :+
  end
end

class Spreadsheet
  PP_CELL_SIZE     = 100
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

  def find_or_create_cell(coord, content = nil)
    (find_cell_coord(coord) || add_cell(coord)).tap do |cell|
      cell.content = content if content
    end
  end

  def set(coord, content)
    find_or_create_cell coord, content
  end

  def add_cell(coord, content = nil)
    raise AlreadyExistentCellError, "Cell #{coord} already exists" if find_cell_coord(coord)

    Cell.new(self, coord, content).tap do |cell|
      update_cell_coord coord, cell
    end
  end

  def move_cell(old_coord, new_coord)
    cell = find_cell_coord(old_coord)

    delete_cell_coord old_coord
    update_cell_coord new_coord, cell
  end

  def add_col(col_to_add, count = 1)
    col_to_add = CellCoordinate.col_coord_index(col_to_add) unless col_to_add.is_a?(Fixnum)

    cells[:by_col].select { |(col, _)| col >= col_to_add }.sort.reverse.each do |(_, rows)|
      rows.sort.each do |(_, cell)|
        cell.move_right! count
      end
    end
  end

  def delete_col(col_to_delete, count = 1)
    col_to_delete = CellCoordinate.col_coord_index(col_to_delete) unless col_to_delete.is_a?(Fixnum)

    cells[:by_col][col_to_delete].each { |(_, cell)| delete_cell_coord cell.coord } if cells[:by_col][col_to_delete]

    cells[:by_col].select { |(col, _)| col >= col_to_delete + count }.sort.each do |(_, rows)|
      rows.sort.each do |(_, cell)|
        cell.move_left! count
      end
    end
  end

  def add_row(row_to_add, count = 1)
    cells[:by_row].select { |(row, _)| row >= row_to_add }.sort.reverse.each do |(_, cols)|
      cols.sort.each do |(_, cell)|
        cell.move_down! count
      end
    end
  end

  def delete_row(row_to_delete, count = 1)
    cells[:by_row][row_to_delete].each { |(_, cell)| delete_cell_coord cell.coord } if cells[:by_row][row_to_delete]

    cells[:by_row].select { |(row, _)| row >= row_to_delete + count }.sort.each do |(_, cols)|
      cols.sort.each do |(_, cell)|
        cell.move_up! count
      end
    end
  end

  def move_col(source_col, dest_col, count = 1)
    source_col = CellCoordinate.col_coord_index(source_col)  unless source_col.is_a?(Fixnum)
    dest_col   = CellCoordinate.col_coord_index(dest_col)    unless dest_col.is_a?(Fixnum)

    if dest_col >= source_col
      return if dest_col - (source_col + count - 1) <= 1

      add_col dest_col, count

      cells[:by_col].select { |(col, _)| col >= source_col && col < source_col + count }.sort.each do |(_, rows)|
        rows.sort.each do |(_, cell)|
          cell.move_right! dest_col - source_col
        end
      end

      delete_col source_col, count
    else
      add_col dest_col, count

      source_col += count

      cells[:by_col].select { |(col, _)| col >= source_col && col < source_col + count }.sort.each do |(_, rows)|
        rows.sort.each do |(_, cell)|
          cell.move_left! source_col - dest_col
        end
      end

      delete_col source_col, count
    end
  end

  def move_row(source_row, dest_row, count = 1)
    if dest_row >= source_row
      return if dest_row - (source_row + count - 1) <= 1

      add_row dest_row, count

      cells[:by_row].select { |(row, _)| row >= source_row && row < source_row + count }.sort.each do |(_, cols)|
        cols.sort.each do |(_, cell)|
          cell.move_down! dest_row - source_row
        end
      end

      delete_row source_row, count
    else
      add_row dest_row, count

      source_row += count

      cells[:by_row].select { |(row, _)| row >= source_row && row < source_row + count }.sort.each do |(_, cols)|
        cols.sort.each do |(_, cell)|
          cell.move_up! source_row - dest_row
        end
      end

      delete_row source_row, count
    end
  end

  def copy_col(source_col, dest_col, count = 1)
    source_col = CellCoordinate.col_coord_index(source_col)  unless source_col.is_a?(Fixnum)
    dest_col   = CellCoordinate.col_coord_index(dest_col)    unless dest_col.is_a?(Fixnum)

    cells[:by_col].select { |(col, _)| col >= source_col && col < source_col + count }.sort.each do |(_, rows)|
      rows.sort.each do |(_, cell)|
        cell.copy_to cell.coord.right_neighbor(dest_col - source_col)
      end
    end
  end

  def copy_row(source_row, dest_row, count = 1)
    source_row = CellCoordinate.row_coord_index(source_row)  unless source_row.is_a?(Fixnum)
    dest_row   = CellCoordinate.row_coord_index(dest_row)    unless dest_row.is_a?(Fixnum)

    cells[:by_row].select { |(row, _)| row >= source_row && row < source_row + count }.sort.each do |(_, cols)|
      cols.sort.each do |(_, cell)|
        cell.copy_to cell.coord.lower_neighbor(dest_row - source_row)
      end
    end
  end

  def consistent?
    cells[:all].all? do |(_, cell)|
      consistent =
        if cell.formula?
          cell_references = cell.find_references

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

      col = cell.coord.col_index
      row = cell.coord.row

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
      puts (1..max_col).map { |col|
        CellCoordinate.col_coord_name(col).to_s.rjust(index = PP_CELL_SIZE / 2 + 1) + ' ' *  (PP_CELL_SIZE - index)
      }.join(PP_COL_DELIMITER)

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

    read_cell_coord = -> (message = 'Enter cell reference: ') do
      coord = read_value.call(message, /^#{CellCoordinate::CELL_COORD}$/i)
    end

    read_cell_range = -> (message = 'Enter cell range: ') do
      coord = read_value.call(message, /^#{CellCoordinate::CELL_RANGE}$/i)
    end

    read_number = -> (message, default_value = nil) do
      number = read_value.call(message, 1..2**32, default_value)
      number.to_i
    end

    last_change = nil

    loop do
      begin
        coord = nil

        pp last_change

        last_change = Time.now

        action = read_value.call(
          "Enter action [S - Set cell (default); M - Move cell; C - Copy cell to cell; CN - Copy cell to range; AC - Add col; AR - Add row; DC - Delete col; DR - Delete row; MC - Move Column; MR - Move Row; CC - Copy col; CR - Copy row; Q - Quit]: ",
          ['S', 'M', 'C', 'CN', 'AR', 'AC', 'DC', 'DR', 'MC', 'MR', 'CC', 'CR', 'Q'],
          'S'
        )

        case action.upcase
        when 'S' then
          coord     = read_cell_coord.call
          content = read_value.call("Enter content (for formulas start with a '='): ")

          set coord, content

        when 'M' then
          subaction = read_value.call(
            'Enter sub action [S - Specific position (default); U - Up; D - Down; L - Left; R - Right]: ',
            ['S', 'U', 'D', 'L', 'R'],
            'S'
          )

          coord = read_cell_coord.call('Select source reference: ')

          cell = find_or_create_cell(coord)

          case subaction.upcase
          when 'S' then
            cell.move_to! read_cell_coord.call('Select destination reference: ')
          when 'U' then
            cell.move_up! read_number.call('Enter # of rows (default: 1): ', 1)
          when 'D' then
            cell.move_down! read_number.call('Enter # of rows (default: 1): ', 1)
          when 'L' then
            cell.move_left! read_number.call('Enter # of cols (default: 1): ', 1)
          when 'R' then
            cell.move_right! read_number.call('Enter # of cols (default: 1): ', 1)
          end

        when 'C' then
          coord      = read_cell_coord.call('Select source reference: ')
          cell     = find_or_create_cell(coord)
          dest_coord = read_cell_coord.call('Select destination reference: ')

          cell.copy_to dest_coord

        when 'CN' then
          coord        = read_cell_coord.call('Select source reference: ')
          cell       = find_or_create_cell(coord)
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

        when 'MC' then
          source_col = read_value.call('Enter source col name (>= "A"): ', 'A'..'ZZZ')
          dest_col   = read_value.call('Enter destination col name (>= "A"): ', 'A'..'ZZZ')
          col_count  = read_number.call('Enter # of cols (default: 1): ', 1)

          move_col source_col, dest_col, col_count

        when 'MR' then
          source_row = read_number.call('Enter source row: ')
          dest_row   = read_number.call('Enter destination row: ')
          row_count  = read_number.call('Enter # of rows (default: 1): ', 1)

          move_row source_row, dest_row, row_count

        when 'CC' then
          source_col = read_value.call('Enter source col name (>= "A"): ', 'A'..'ZZZ')
          dest_col   = read_value.call('Enter destination col name (>= "A"): ', 'A'..'ZZZ')
          col_count  = read_number.call('Enter # of cols (default: 1): ', 1)

          copy_col source_col, dest_col, col_count

        when 'CR' then
          source_row = read_number.call('Enter source row: ')
          dest_row   = read_number.call('Enter destination row: ')
          row_count  = read_number.call('Enter # of rows (default: 1): ', 1)

          copy_row source_row, dest_row, row_count

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

      consistent?
    end
  end

  private

  def find_cell_coord(coord)
    coord = CellCoordinate.self_or_new(coord)

    cells[:all][coord.coord]
  end

  def delete_cell_coord(coord)
    coord = CellCoordinate.self_or_new(coord)

    col = coord.col_index
    row = coord.row

    cells[:by_col][col] ||= {}
    cells[:by_row][row] ||= {}

    cells[:all].delete coord.coord
    cells[:by_col][col].delete row
    cells[:by_row][row].delete col
  end

  def update_cell_coord(coord, cell)
    coord = CellCoordinate.self_or_new(coord)

    col = coord.col_index
    row = coord.row

    cells[:by_col][col] ||= {}
    cells[:by_row][row] ||= {}

    cells[:all][coord.coord] = cells[:by_col][col][row] = cells[:by_row][row][col] = cell
  end
end

def run!
  spreadsheet = Spreadsheet.new

  # # Fibonacci sequence.
  # b1 = spreadsheet.set(:B1, 'Fibonacci sequence:')
  # a3 = spreadsheet.set(:A3, 1)
  # a4 = spreadsheet.set(:A4, '=A3+1')
  # a4.copy_to_range 'A5:A20'
  # b3 = spreadsheet.set(:B3, 1)
  # b4 = spreadsheet.set(:B4, 1)
  # b5 = spreadsheet.set(:B5, '=B3+B4')
  # b5.copy_to_range 'B6:B20'
  #
  # # Factorials.
  # e1 = spreadsheet.set(:E1, 'Factorials:')
  # d3 = spreadsheet.set(:D3, 1)
  # e3 = spreadsheet.set(:E3, '=D3')
  # d4 = spreadsheet.set(:D4, '=D3+1')
  # e4 = spreadsheet.set(:E4, '=D4*E3')
  # d4.copy_to_range 'D5:D20'
  # e4.copy_to_range 'E5:E20'

  # Case with performance problems.
  a1 = spreadsheet.set(:A1, 1)
  a2 = spreadsheet.set(:A2, '=A1+1')
  a2.copy_to_range 'A3:A1000'
  spreadsheet.set(:A1001, '=sum(A1:A1000)')

  spreadsheet.repl
end

run! if __FILE__ == $0
