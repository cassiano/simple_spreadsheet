require 'colorize'
require 'after_do'

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
      if (truncate_position = limit - 1 - delimiter.size) >= 0
        self[0..truncate_position] + delimiter
      elsif (delimiter_truncate_position = limit - 1) >= 0
        delimiter[0..delimiter_truncate_position]
      else
        ''
      end
    end
  end
end

class CellAddress
  COL_RANGE                      = ('A'..'ZZZ').to_a.map(&:to_sym)
  CELL_COORD_UPPER_CASE          = '[A-Z]+[1-9]\d*'
  CELL_COORD_LOWER_CASE          = '[a-z]+[1-9]\d*'
  CELL_COORD                     = '\$?[A-Z]+\$?[1-9]\d*'
  CELL_COORD_WITH_PARENS         = '(\$?)([A-Z]+)(\$?)([1-9]\d*)'
  CELL_COORD_REG_EXP             = /#{CELL_COORD}/i
  CELL_COORD_WITH_PARENS_REG_EXP = /#{CELL_COORD_WITH_PARENS}/i
  CELL_RANGE                     = "#{CELL_COORD_UPPER_CASE}:#{CELL_COORD_UPPER_CASE}"
  CELL_RANGE_WITH_PARENS         = "(#{CELL_COORD_UPPER_CASE}):(#{CELL_COORD_UPPER_CASE})"
  CELL_RANGE_WITH_PARENS_REG_EXP = /(#{CELL_RANGE_WITH_PARENS})/i

  # List of possible exceptions.
  class IllegalCellReference < StandardError; end

  attr_reader :addr

  delegate :col_addr_index, :col_addr_name, :normalize_addr, :parse_addr, to: :class

  def initialize(*addr)
    addr.flatten!
    addr[0] = col_addr_name(addr[0]) if addr[0].is_a?(Fixnum)
    addr    = addr.join

    raise IllegalCellReference, "Illegal cell reference `#{addr}`" unless addr =~ /^#{CellAddress::CELL_COORD}$/i

    @addr = normalize_addr(addr)
  end

  def self.self_or_new(addr_or_cell_addr)
    if addr_or_cell_addr.is_a?(CellAddress)
      addr_or_cell_addr
    else
      new addr_or_cell_addr
    end
  end

  def col
    @col ||= col_and_row[0]
  end

  def col_index
    @col_index ||= col_addr_index(col)
  end

  def row
    @row ||= col_and_row[1]
  end

  def col_and_row
    @col_and_row ||= parse_addr(addr)
  end

  def neighbor(col_count: 0, row_count: 0)
    raise IllegalCellReference unless col_index + col_count > 0 && row + row_count > 0

    CellAddress.new col_addr_name(col_index + col_count), row + row_count
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
    addr.to_sym
  end

  def to_s
    addr.upcase.to_s
  end

  def ==(other_addr)
    addr == (other_addr.is_a?(CellAddress) ? other_addr.addr : normalize_addr(other_addr))
  end

  def self.normalize_addr(addr)
    parse_addr(addr).join.to_sym
  end

  def self.parse_addr(addr)
    addr.to_s =~ CELL_COORD_WITH_PARENS_REG_EXP && [$2.upcase.to_sym, $4.to_i]
  end

  def self.col_addr_index(col_addr)
    COL_RANGE.index(col_addr.upcase.to_sym) + 1
  end

  def self.col_addr_name(col_index)
    COL_RANGE[col_index - 1]
  end

  def self.splat_range(upper_left_addr, lower_right_addr)
    upper_left_addr  = new(upper_left_addr)  unless upper_left_addr.is_a?(self)
    lower_right_addr = new(lower_right_addr) unless lower_right_addr.is_a?(self)

    ul_col, ul_row = upper_left_addr.col_and_row
    lr_col, lr_row = lower_right_addr.col_and_row

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

  attr_reader   :spreadsheet, :addr, :references, :observers, :content, :evaluatable_content, :last_evaluated_at, :eval_error
  attr_accessor :max_reference_timestamp

  def initialize(spreadsheet, addr, content = nil)
    log "Creating cell #{addr}"

    addr = CellAddress.self_or_new(addr)

    @spreadsheet             = spreadsheet
    @addr                    = addr
    @references              = []
    @observers               = []
    @max_reference_timestamp = nil
    @eval_error              = nil

    self.content = content
  end

  def add_observer(observer)
    log "Adding observer #{observer.addr} to #{addr}"

    observers.unique_add observer
  end

  def remove_observer(observer)
    log "Removing observer #{observer.addr} from #{addr}"

    observers.delete observer
  end

  def content=(new_content)
    log "Replacing content `#{content}` with new content `#{new_content}` in cell #{addr}"

    new_content.strip! if new_content.is_a?(String)

    if new_content.is_a?(String)
      # Upcase cell addresses (e.g. 'a1' by 'A1').
      new_content.gsub!(/(#{CellAddress::CELL_COORD})/i) { $1.upcase }

      # Downcase formula function names (e.g. 'SUM(A1: A2)' by 'sum(A1:A2)').
      Formula.singleton_methods.each { |method| new_content.gsub! /\b#{method}\(/i, "#{method}(" }

      @content, @evaluatable_content = new_content.clone, new_content.clone
    else
      @content, @evaluatable_content = new_content, new_content
    end

    @last_evaluated_at = nil    # Assume it was never evaluated before, since its contents have effectively changed.

    if formula?
      # Splat ranges, e.g., replace 'A1:A3' by '[[A1, A2, A3]]'.
      @evaluatable_content.scan(CellAddress::CELL_RANGE_WITH_PARENS_REG_EXP).uniq.each do |(range, upper_left_addr, lower_right_addr)|
        upper_left_col, upper_left_row   = CellAddress.parse_addr(upper_left_addr)
        lower_right_col, lower_right_row = CellAddress.parse_addr(lower_right_addr)
        upper_left_col                   = CellAddress.col_addr_index(upper_left_col)
        lower_right_col                  = CellAddress.col_addr_index(lower_right_col)

        @evaluatable_content.gsub!  /(?<![A-Z])#{Regexp.escape(range)}(?![0-9])/i,
                                    '[' + CellAddress.splat_range(upper_left_addr, lower_right_addr).map { |row|
                                      '[' + row.map { |addr| addr.addr.downcase }.join(',') + ']'
                                    }.join(', ') + '],' + {
                                      upper_left: [upper_left_col, upper_left_row],
                                      lower_right: [lower_right_col, lower_right_row]
                                    }.to_s
      end

      # Replace all cell addresses by template variables (e.g. 'A1' by '%{A1}', '$A1' by '%{$A1}' etc).
      @evaluatable_content.gsub! /(#{CellAddress::CELL_COORD})/i, '%{\1}'

      log "Evaluatable content before finding references: `#{@evaluatable_content}`"
    end

    eval true
  end

  def eval(reevaluate = false)
    previous_content = @evaluated_content

    if reevaluate
      if !formula? || !(max_reference_timestamp && last_evaluated_at && last_evaluated_at >= max_reference_timestamp)
        @evaluated_content = nil    # Turn cache off.
      else
        log "Skiping reevaluation for #{addr}"
      end
    end

    @evaluated_content ||= begin
      @eval_error = nil

      begin
        sync_references
      rescue StandardError => e
        @eval_error = e
      end

      new_evaluated_content =
        if invalid?
          eval_error_message
        elsif formula?
          calculate_formula
        else    # Scalar.
          content
        end

      @last_evaluated_at = Time.now

      new_evaluated_content
    end

    # Fire all observers if evaluated content has changed.
    fire_observers if previous_content != @evaluated_content

    @evaluated_content || DEFAULT_VALUE
  end

  def find_references
    return [] unless formula?

    evaluatable_content.scan(/%\{(#{CellAddress::CELL_COORD})\}/i).uniq.inject [] do |memo, (addr)|
      cell           = spreadsheet.find_or_create_cell(addr)
      cell_reference = CellReference.new(
        cell,
        addr: addr,
        is_range: !!(addr =~ Regexp.new(CellAddress::CELL_COORD_LOWER_CASE))
      )

      memo.unique_add cell_reference
    end
  end

  def valid?
    eval_error.nil?
  end

  def invalid?
    !valid?
  end

  def eval_error_message
    @eval_error && @eval_error.message
  end

  def blank?
    @evaluated_content.nil?
  end

  def reset_circular_reference_check_cache
    return unless @directly_or_indirectly_references

    log "Resetting circular reference check cache for #{addr}"

    @directly_or_indirectly_references = nil

    reset_observers_circular_reference_check_cache
  end

  def reset_observers_circular_reference_check_cache
    observers.each &:reset_circular_reference_check_cache
  end

  def directly_or_indirectly_references?(cell)
    log "Checking if #{cell.addr} directly or indirectly references #{addr}"

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
    CellAddress.splat_range(*dest_range.split(':')).flatten.each do |addr|
      copy_to addr
    end
  end

  def copy_to(dest_addr)
    dest_addr = CellAddress.self_or_new(dest_addr)

    return if dest_addr == addr

    if content.is_a?(String)
      dest_content = content.clone

      log "Content before replacements in copy_to: #{dest_content}"

      dest_content.gsub! CellAddress::CELL_COORD_WITH_PARENS_REG_EXP do |content_addr|
        if (cell = references.find { |reference| reference == content_addr })
          cell.new_addr addr, dest_addr
        end
      end

      log "Content after replacements in copy_to: #{dest_content}"
    else
      dest_content = content
    end

    spreadsheet.set dest_addr, dest_content
  end

  def move_to!(dest_addr, update_ranges_mode: false, direction: nil, affected_cols: nil, affected_rows: nil)
    log "Moving #{addr} to #{dest_addr}"

    dest_addr = CellAddress.self_or_new(dest_addr)

    return if dest_addr == addr

    if formula?
      if update_ranges_mode
        new_content = content.clone

        content.scan(CellAddress::CELL_RANGE_WITH_PARENS_REG_EXP).uniq.each do |(range, upper_left_addr, lower_right_addr)|
          upper_left_col, upper_left_row   = CellAddress.parse_addr(upper_left_addr)
          lower_right_col, lower_right_row = CellAddress.parse_addr(lower_right_addr)

          update_upper_left_cell =
            case direction
            when :left, :right
              upper_left_col = CellAddress.col_addr_index(upper_left_col)

              affected_cols.include? upper_left_col
            when :up, :down
              affected_rows.include? upper_left_row
            end

          update_lower_right_cell =
            case direction
            when :left, :right
              lower_right_col = CellAddress.col_addr_index(lower_right_col)

              affected_cols.include? lower_right_col
            when :up, :down
              affected_rows.include? lower_right_row
            end

          upper_left_cell  = CellReference.new(spreadsheet.find_or_create_cell(upper_left_addr),  addr: upper_left_addr)
          lower_right_cell = CellReference.new(spreadsheet.find_or_create_cell(lower_right_addr), addr: lower_right_addr)

          upper_left_cell_new_addr  = update_upper_left_cell  ? upper_left_cell.new_addr(addr, dest_addr)  : upper_left_cell.addr
          lower_right_cell_new_addr = update_lower_right_cell ? lower_right_cell.new_addr(addr, dest_addr) : lower_right_cell.addr

          log "Updating range `#{range}` to `#{upper_left_cell_new_addr}:#{lower_right_cell_new_addr}`"

          new_content.gsub! /(?<![A-Z])#{range}(?![0-9])/i, [upper_left_cell_new_addr, lower_right_cell_new_addr].join(':')
        end
      else
        new_content = content.clone
      end
    else
      new_content = content
    end

    spreadsheet.set dest_addr, new_content

    log "#{addr}'s observers: #{observers.map { |o| o.addr.addr }.join(', ')}"

    # Update all non-range observers so they refer to the new address.
    observers.reject(&:range?).each do |observer|
      observer.update_reference addr, dest_addr
    end

    # Reset the current cell's value, with the added value that it will automatically reevaluate all remaining (range) observers.
    self.content = nil
  end

  def move_right!(col_count = 1, update_ranges_mode: false, affected_cols: nil)
    move_to! addr.right_neighbor(col_count), update_ranges_mode: update_ranges_mode, direction: :right, affected_cols: affected_cols
  end

  def move_left!(col_count = 1, update_ranges_mode: false, affected_cols: nil)
    move_to! addr.left_neighbor(col_count), update_ranges_mode: update_ranges_mode, direction: :left, affected_cols: affected_cols
  end

  def move_down!(row_count = 1, update_ranges_mode: false, affected_rows: nil)
    move_to! addr.lower_neighbor(row_count), update_ranges_mode: update_ranges_mode, direction: :down, affected_rows: affected_rows
  end

  def move_up!(row_count = 1, update_ranges_mode: false, affected_rows: nil)
    move_to! addr.upper_neighbor(row_count), update_ranges_mode: update_ranges_mode, direction: :up, affected_rows: affected_rows
  end

  def update_reference(old_addr, new_addr)
    log "Replacing reference `#{old_addr}` with `#{new_addr}` in #{addr}"

    old_col, old_row = CellAddress.parse_addr(old_addr)
    new_col, new_row = CellAddress.parse_addr(new_addr)

    # Do not use gsub! (since the setter won't be called).
    self.content = content.gsub(/(?<![A-Z:])(\$?)#{old_col}(\$?)#{old_row}(?![0-9:])/i) { [$1, new_col, $2, new_row].join }
  end

  def formula?
    content.is_a?(String) && content[0] == '='
  end

  def scalar?
    !formula?
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
      address:     addr.addr,
      references:  references,
      observers:   observers
    }
  end

  private

  def calculate_formula
    log "Calculating formula for #{addr}"

    evaluated_content = evaluatable_content[1..-1]

    # Remove all address (absolute) markers ('$') before evaluating the formula.
    evaluated_content.gsub! /#{CellAddress::CELL_COORD_WITH_PARENS}/i, '\2\4'

    # Evaluate the cell only if all references are valid.
    if valid? && references.all?(&:valid?)
      references_hash = references.inject Hash.new do |memo, reference|
        reference_value = reference.eval.to_s

        memo[reference.addr.addr]           = reference_value
        memo[reference.addr.addr.downcase]  = reference_value

        memo
      end

      log "References hash: #{references_hash.inspect}"

      # Replace all references by their corresponding values and evaluate the cell's content in the Formula context, so functions
      # like `sum`, `average` etc are simply treated as calls to Formula's (singleton) methods.
      begin
        Formula.instance_eval { eval evaluated_content % references_hash }.tap do |value|
          log "Formula value for #{addr}: #{value}"
        end
      rescue StandardError => e
        @eval_error = e
        eval_error_message
      end
    else
      @eval_error ||= references.find(&:invalid?).eval_error
      eval_error_message
    end
  end

  def sync_references
    old_references = references.clone
    new_references = find_references

    references_to_add    = new_references.subtract(old_references)   # Do not use `new_references - old_references`
    references_to_remove = old_references.subtract(new_references)   # and `old_references - new_references`.

    if references_to_add.any? || references_to_remove.any?
      # Notify all direct and indirect observers to reset their circular reference check cache (memoization). Notice
      # that this step must be done necessarily BEFORE adding or removing references, given observers
      reset_circular_reference_check_cache

      add_references    references_to_add
      remove_references references_to_remove
    end
  end

  def fire_observers
    log "Firing #{addr}'s observers" if observers.any?

    observers.each do |observer|
      if last_evaluated_at && (!observer.max_reference_timestamp || last_evaluated_at > observer.max_reference_timestamp)
        observer.max_reference_timestamp = last_evaluated_at
      end

      observer.eval true
    end
  end

  def add_reference(reference)
    log "Adding reference #{reference.addr} to #{addr} (range? #{reference.range?})"

    if reference.directly_or_indirectly_references?(self)
      raise CircularReferenceError, "Circular reference (#{reference.addr} <- #{addr})"
    end

    references.unique_add reference
    reference.add_observer CellReference.new(self, is_range: reference.range?)

    if reference.last_evaluated_at && (!max_reference_timestamp || reference.last_evaluated_at > max_reference_timestamp)
      self.max_reference_timestamp = reference.last_evaluated_at
    end
  end

  def remove_all_references
    self.max_reference_timestamp = nil

    references.each do |reference|
      remove_reference reference
    end
  end

  def remove_reference(reference)
    log "Removing reference #{reference.addr} from #{addr}"

    references.delete reference
    reference.remove_observer(self) unless references.any? { |addr| addr.cell == reference.cell }

    if reference.last_evaluated_at && max_reference_timestamp && reference.last_evaluated_at == max_reference_timestamp
      self.max_reference_timestamp = find_max_reference_timestamp
    end

    # Physically remove reference if blank and no other cells point to it.
    spreadsheet.delete_cell_addr(reference.addr) if reference.blank? && reference.observers.empty?
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

  def find_max_reference_timestamp
    references.map(&:last_evaluated_at).compact.max
  end
end

class CellReference
  attr_reader :cell

  delegate_all to: :cell

  def initialize(cell, addr: cell.addr.addr, is_range: false)
    raise IllegalCellReference unless addr.to_s =~ CellAddress::CELL_COORD_WITH_PARENS_REG_EXP

    @cell            = cell
    @is_absolute_col = $1 == '$'
    @is_absolute_row = $3 == '$'
    @is_range   = is_range
  end

  def absolute_col?; @is_absolute_col; end
  def absolute_row?; @is_absolute_row; end
  def range?;   @is_range;   end

  def full_addr
    col, row = cell.addr.col_and_row

    addr_parts = []
    addr_parts << '$' if absolute_col?
    addr_parts << col
    addr_parts << '$' if absolute_row?
    addr_parts << row

    addr_parts.join
  end

  # Calculates a cell's new reference when an observer cell is copied from `observer_source_addr` to `observer_dest_addr`.
  def new_addr(observer_source_addr, observer_dest_addr)
    col_diff = absolute_col? ? 0 : observer_dest_addr.col_index - observer_source_addr.col_index
    row_diff = absolute_row? ? 0 : observer_dest_addr.row       - observer_source_addr.row

    target_addr = addr.right_neighbor(col_diff).lower_neighbor(row_diff)

    addr_parts = []
    addr_parts << '$' if absolute_col?
    addr_parts << target_addr.col
    addr_parts << '$' if absolute_row?
    addr_parts << target_addr.row

    addr_parts.join
  end

  def ==(another_cell_or_cell_reference)
    case another_cell_or_cell_reference
    when CellReference then
      cell == another_cell_or_cell_reference.cell &&
        absolute_col? == another_cell_or_cell_reference.absolute_col? &&
        absolute_row? == another_cell_or_cell_reference.absolute_row? &&
        range?   == another_cell_or_cell_reference.range?
    when Cell then
      cell == another_cell_or_cell_reference
    when String, Symbol then
      full_addr == another_cell_or_cell_reference.to_s.upcase
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
  #     [any_object.to_s.upcase, full_addr]
  #   else
  #     super
  #   end
  # end
end

class Formula
  def self.sum(*cell_values)
    @cell_values.flatten.inject :+
  end

  def self.count(*cell_values)
    @cell_values.flatten.size
  end

  def self.average(*cell_values)
    sum(@cell_values) * 1.0 / count(@cell_values)
  end

  def self.max(*cell_values)
    @cell_values.flatten.max
  end

  def self.min(*cell_values)
    @cell_values.flatten.min
  end

  def self.col_count(*cell_values)
    @cell_values[0][0].size   # The size of all rows should be the same, so we can safely check the 1st row.
  end

  def self.row_count(*cell_values)
    @cell_values[0].size
  end

  def self.col_num(*cell_values)
    @range_metadata && @range_metadata[:upper_left][0]
  end

  def self.row_num(*cell_values)
    @range_metadata && @range_metadata[:upper_left][1]
  end

  singleton_class.extend AfterDo    # https://github.com/PragTob/after_do

  singleton_class.before singleton_methods do |*cell_values|
    # Get rid of the last argument, usually the Formula class itself.
    cell_values.pop if cell_values.last == self

    @range_metadata = cell_values.pop if cell_values.last.is_a?(Hash)
    @cell_values    = cell_values
  end
end

class Spreadsheet
  PP_CELL_SIZE     = 30
  PP_ROW_REF_SIZE  = 5
  PP_COL_DELIMITER = ' | '
  PP_EMPTY_CELL    = '·'

  # List of possible exceptions.
  class AlreadyExistentCellError < StandardError; end

  attr_reader :cells

  def initialize
    @cells = {
      all: {},
      cols: {},
      rows: {}
    }
  end

  def cell_count
    cells[:all].values.compact.size
  end

  def find_or_create_cell(addr, content = nil)
    (find_cell_addr(addr) || add_cell(addr)).tap do |cell|
      cell.content = content if content
    end
  end

  def set(addr, content)
    find_or_create_cell addr, content
  end

  def add_cell(addr, content = nil)
    raise AlreadyExistentCellError, "Cell #{addr} already exists" if find_cell_addr(addr)

    Cell.new(self, addr, content).tap do |cell|
      update_cell_addr addr, cell
    end
  end

  def move_cell(old_addr, new_addr)
    cell = find_cell_addr(old_addr)

    delete_cell_addr old_addr
    update_cell_addr new_addr, cell
  end

  def add_col(col_to_add, count = 1)
    col_to_add = CellAddress.col_addr_index(col_to_add) unless col_to_add.is_a?(Fixnum)

    affected_cols = col_to_add..last_col_index

    cells[:cols].select { |(col, _)| col >= col_to_add }.sort.reverse.each do |(_, rows)|
      rows.sort.each do |(_, cell)|
        cell.move_right! count, update_ranges_mode: true, affected_cols: affected_cols
      end
    end
  end

  def delete_col(col_to_delete, count = 1)
    col_to_delete = CellAddress.col_addr_index(col_to_delete) unless col_to_delete.is_a?(Fixnum)

    affected_cols = (col_to_delete + count)..last_col_index

    count.times do |i|
      cells[:cols][col_to_delete + i].each { |(_, cell)| delete_cell_addr cell.addr } if cells[:cols][col_to_delete + i]
    end

    cells[:cols].select { |(col, _)| col >= col_to_delete + count }.sort.each do |(_, rows)|
      rows.sort.each do |(_, cell)|
        cell.move_left! count, update_ranges_mode: true, affected_cols: affected_cols
      end
    end
  end

  def add_row(row_to_add, count = 1)
    affected_rows = row_to_add..last_row

    cells[:rows].select { |(row, _)| row >= row_to_add }.sort.reverse.each do |(_, cols)|
      cols.sort.each do |(_, cell)|
        cell.move_down! count, update_ranges_mode: true, affected_rows: affected_rows
      end
    end
  end

  def delete_row(row_to_delete, count = 1)
    affected_rows = (row_to_delete + count)..last_row

    count.times do |i|
      cells[:rows][row_to_delete + i].each { |(_, cell)| delete_cell_addr cell.addr } if cells[:rows][row_to_delete + i]
    end

    cells[:rows].select { |(row, _)| row >= row_to_delete + count }.sort.each do |(_, cols)|
      cols.sort.each do |(_, cell)|
        cell.move_up! count, update_ranges_mode: true, affected_rows: affected_rows
      end
    end
  end

  def move_col(source_col, dest_col, count = 1)
    source_col = CellAddress.col_addr_index(source_col)  unless source_col.is_a?(Fixnum)
    dest_col   = CellAddress.col_addr_index(dest_col)    unless dest_col.is_a?(Fixnum)

    if dest_col >= source_col
      return if dest_col - (source_col + count - 1) <= 1

      add_col dest_col, count

      cells[:cols].select { |(col, _)| col >= source_col && col < source_col + count }.sort.each do |(_, rows)|
        rows.sort.each do |(_, cell)|
          cell.move_right! dest_col - source_col, update_ranges_mode: true, affected_cols: source_col..(source_col + count - 1)
        end
      end

      delete_col source_col, count
    else
      add_col dest_col, count

      source_col += count

      cells[:cols].select { |(col, _)| col >= source_col && col < source_col + count }.sort.each do |(_, rows)|
        rows.sort.each do |(_, cell)|
          cell.move_left! source_col - dest_col, update_ranges_mode: true, affected_cols: source_col..(source_col + count - 1)
        end
      end

      delete_col source_col, count
    end
  end

  def move_row(source_row, dest_row, count = 1)
    if dest_row >= source_row
      return if dest_row - (source_row + count - 1) <= 1

      add_row dest_row, count

      cells[:rows].select { |(row, _)| row >= source_row && row < source_row + count }.sort.each do |(_, cols)|
        cols.sort.each do |(_, cell)|
          cell.move_down! dest_row - source_row, update_ranges_mode: true, affected_rows: source_row..(source_row + count - 1)
        end
      end

      delete_row source_row, count
    else
      add_row dest_row, count

      source_row += count

      cells[:rows].select { |(row, _)| row >= source_row && row < source_row + count }.sort.each do |(_, cols)|
        cols.sort.each do |(_, cell)|
          cell.move_up! source_row - dest_row, update_ranges_mode: true, affected_rows: source_row..(source_row + count - 1)
        end
      end

      delete_row source_row, count
    end
  end

  def copy_col(source_col, dest_col, count = 1)
    source_col = CellAddress.col_addr_index(source_col)  unless source_col.is_a?(Fixnum)
    dest_col   = CellAddress.col_addr_index(dest_col)    unless dest_col.is_a?(Fixnum)

    cells[:cols].select { |(col, _)| col >= source_col && col < source_col + count }.sort.each do |(_, rows)|
      rows.sort.each do |(_, cell)|
        cell.copy_to cell.addr.right_neighbor(dest_col - source_col)
      end
    end
  end

  def copy_row(source_row, dest_row, count = 1)
    source_row = CellAddress.row_addr_index(source_row)  unless source_row.is_a?(Fixnum)
    dest_row   = CellAddress.row_addr_index(dest_row)    unless dest_row.is_a?(Fixnum)

    cells[:rows].select { |(row, _)| row >= source_row && row < source_row + count }.sort.each do |(_, cols)|
      cols.sort.each do |(_, cell)|
        cell.copy_to cell.addr.lower_neighbor(dest_row - source_row)
      end
    end
  end

  def consistent?
    cells[:all].all? do |(_, cell)|
      log "Checking #{cell.addr} consistency"

      log "Checking references and corresponding observers"

      if cell.valid?
        consistent =
          if cell.formula?
            cell_references = cell.find_references

            log "Checking if references found (#{cell_references.map(&:addr).join(', ')}) correspond " +
                "to stored ones (#{cell.references.map(&:addr).join(', ')})"

            cell_references.size == cell.references.size && (cell_references.map(&:addr) - cell.references.map(&:addr)).empty? && cell.references.all? do |reference|
              log "Checking if #{reference.addr} reference's observers include it"

              reference.observers.include? cell
            end
          else
            log "Checking if stored references is empty"

            cell.references.empty?
          end

        log "Error(s) detected" unless consistent

        next false unless consistent

        log "Checking observers"

        consistent = cell.observers.all? do |observer|
          observer.references.include? cell
        end

        log "Error(s) detected" unless consistent

        next false unless consistent

        log "Checking spreadsheet data structures"

        col = cell.addr.col_index
        row = cell.addr.row

        consistent =
          cells[:cols][col] && cells[:cols][col][row] == cell &&
          cells[:rows][row] && cells[:rows][row][col] == cell

        log "Error(s) detected" unless consistent

        next false unless consistent
      else
        log "Bypassing further checkings for this cell, because it is invalid (error message: `#{cell.eval_error}`)."
      end

      true
    end
  end

  def pp(last_change)
    lrjust = -> (ltext, rtext, size) do
      spacing_size = size - (ltext.size + rtext.size)

      if spacing_size >= 0
        ltext + ' ' * spacing_size + rtext
      elsif rtext.size < size
        lrjust.call ltext.truncate(ltext.size + spacing_size - 1), rtext, size
      else
        rtext.truncate size
      end
    end

    max_col = (max = cells[:cols].sort.max) && max[0]
    max_row = (max = cells[:rows].sort.max) && max[0]

    if max_col && max_row
      print ' '
      print ' ' * PP_ROW_REF_SIZE
      puts (1..max_col).map { |col|
        CellAddress.col_addr_name(col).to_s.rjust(index = PP_CELL_SIZE / 2 + 1) + ' ' *  (PP_CELL_SIZE - index)
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

          if (cell = cells[:rows][row] && cells[:rows][row][col])
            value = cell.blank? ? PP_EMPTY_CELL : cell.eval

            highlight_cell = false

            text =
              if cell.formula?
                # Highlight cell if value has changed.
                highlight_cell = last_change && cell.last_evaluated_at && cell.last_evaluated_at > last_change

                lrjust.call("`#{cell.content}`", value.to_s, PP_CELL_SIZE)
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

    read_cell_addr = -> (message = 'Enter cell reference: ') do
      addr = read_value.call(message, /^#{CellAddress::CELL_COORD}$/i)
    end

    read_cell_range = -> (message = 'Enter cell range: ') do
      addr = read_value.call(message, /^#{CellAddress::CELL_RANGE}$/i)
    end

    read_number = -> (message, default_value = nil) do
      number = read_value.call(message, 1..2**32, default_value)
      number.to_i
    end

    last_change = nil

    loop do
      begin
        addr = nil

        pp last_change

        last_change = Time.now

        action = read_value.call(
          "Enter action [S - Set cell (default); M - Move cell; C - Copy cell to cell; CN - Copy cell to range; AC - Add col; AR - Add row; DC - Delete col; DR - Delete row; MC - Move Column; MR - Move Row; CC - Copy col; CR - Copy row; SZ - Set cell size; Q - Quit]: ",
          ['S', 'M', 'C', 'CN', 'AR', 'AC', 'DC', 'DR', 'MC', 'MR', 'CC', 'CR', 'SZ', 'Q'],
          'S'
        )

        case action.upcase
        when 'S' then
          addr     = read_cell_addr.call
          content = read_value.call("Enter content (for formulas start with a '='): ")

          set addr, content

        when 'M' then
          addr = read_cell_addr.call('Select source reference: ')

          subaction = read_value.call(
            'Enter sub action [S - Specific position (default); U - Up; D - Down; L - Left; R - Right]: ',
            ['S', 'U', 'D', 'L', 'R'],
            'S'
          )

          cell = find_or_create_cell(addr)

          case subaction.upcase
          when 'S' then
            cell.move_to! read_cell_addr.call('Select destination reference: ')
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
          addr      = read_cell_addr.call('Select source reference: ')
          cell     = find_or_create_cell(addr)
          dest_addr = read_cell_addr.call('Select destination reference: ')

          cell.copy_to dest_addr

        when 'CN' then
          addr        = read_cell_addr.call('Select source reference: ')
          cell       = find_or_create_cell(addr)
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

        when 'SZ' then
          Spreadsheet.const_set :PP_CELL_SIZE,
                                read_number.call("Enter new cell size (default: #{Spreadsheet::PP_CELL_SIZE}): ", Spreadsheet::PP_CELL_SIZE)

        when 'Q' then
          exit

        else
          next
        end

      rescue StandardError => e
        puts "Error: `#{e}`."
        puts 'Stack trace:'
        puts e.backtrace
        exit
      end

      unless consistent?
        puts "Sorry, but the spreadsheet became inconsistent! Exiting now..."
        exit
      end
    end
  end

  def delete_cell_addr(addr)
    addr = CellAddress.self_or_new(addr)

    col = addr.col_index
    row = addr.row

    cells[:cols][col] ||= {}
    cells[:rows][row] ||= {}

    cells[:all].delete addr.addr
    cells[:cols][col].delete row
    cells[:rows][row].delete col
  end

  private

  def find_cell_addr(addr)
    addr = CellAddress.self_or_new(addr)

    cells[:all][addr.addr]
  end

  def update_cell_addr(addr, cell)
    addr = CellAddress.self_or_new(addr)

    col = addr.col_index
    row = addr.row

    cells[:cols][col] ||= {}
    cells[:rows][row] ||= {}

    cells[:all][addr.addr] = cells[:cols][col][row] = cells[:rows][row][col] = cell
  end

  def last_col_index
    cells[:cols].keys.max
  end

  def last_row
    cells[:rows].keys.max
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
  # c1 = spreadsheet.set(:C1, 'Factorials:')
  # c3 = spreadsheet.set(:C3, 1)
  # c4 = spreadsheet.set(:C4, '=A4*C3')
  # c4.copy_to_range 'C5:C20'

  spreadsheet.set :A1, 1
  a2 = spreadsheet.set(:A2, '=A1+1')
  last_row = 10
  a2.copy_to_range "A3:A#{last_row}"
  spreadsheet.set [:A, last_row + 1], "=sum(A1:A#{last_row})"
  spreadsheet.set [:A, last_row + 2], "=average(A1:A#{last_row})"
  spreadsheet.set [:A, last_row + 3], "=count(A1:A#{last_row})"
  spreadsheet.set [:A, last_row + 4], "=min(A1:A#{last_row})"
  spreadsheet.set [:A, last_row + 5], "=max(A1:A#{last_row})"
  spreadsheet.set [:A, last_row + 6], "=col_count(A1:D#{last_row})"
  spreadsheet.set [:A, last_row + 7], "=row_count(A1:D#{last_row})"
  4.times do |i|
    col_name = CellAddress.col_addr_name(1 + i)
    spreadsheet.set [col_name, last_row + 8], "=col_num(#{col_name}1:#{col_name}#{last_row})"
  end
  4.times do |i|
    spreadsheet.set [:A, last_row + 9 + i], "=row_num(B#{last_row + 9 + i}:D#{last_row + 9 + i})"
  end

  spreadsheet.repl
end

run! if __FILE__ == $0
