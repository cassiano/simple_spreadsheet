DEBUG = true

require 'set'

class Spreadsheet
  attr_reader :cells

  def initialize
    @cells = {}
  end

  # def set(ref, content = nil)
  #   cell = find_or_create_cell_by_ref(ref)
  #
  #   cell.content = content
  # end

  def find_or_create_cell_by_ref(ref)
    ref = ref.upcase.to_sym

    cells[ref] || add_cell(ref)
  end

  def add_cell(ref, content = nil)
    ref = ref.upcase.to_sym

    Cell.new(self, ref, content).tap do |cell|
      cells[ref] = cell
    end
  end
end

class Cell
  CELL_REF_REG_EXP = /[A-Z]+\d+/i
  DEFAULT_VALUE    = 0

  attr_reader :spreadsheet, :ref, :references, :observers, :content

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
    old_references = Set.new
    new_references = Set.new

    old_references = references.clone if is_formula?

    # Notice this may change the value returned by method `is_formula?`.
    @content = String === new_content ? new_content.strip : new_content

    if is_formula?
      new_references = content[1..-1].scan(CELL_REF_REG_EXP).inject Set.new do |memo, ref|
        memo << spreadsheet.find_or_create_cell_by_ref(ref)
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

        evaluated_content = content[1..-1]

        references.each do |cell|
          evaluated_content.gsub! cell.ref.to_s, cell.eval.to_s
        end

        Kernel.eval evaluated_content
      else
        content
      end

    # Fire all observers if evaluated content has changed.
    fire_observers if previous_content != @evaluated_content

    @evaluated_content || DEFAULT_VALUE
  end

  def direct_or_indirect_references
    cells = references.inject(references.clone) do |memo, cell|
      memo << cell.direct_or_indirect_references
    end

    cells.flatten
  end

  private

  def fire_observers
    puts "Firing #{ref}'s observers" if DEBUG && observers.any?

    observers.each do |cell|
      cell.eval true
    end
  end

  def add_reference(cell)
    raise "Cyclical reference detected when adding reference #{cell.ref} to #{ref}!" if cell.direct_or_indirect_references.include?(self)

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
end

spreadsheet = Spreadsheet.new

# a1 = spreadsheet.add_cell :A1, 10
# b1 = spreadsheet.add_cell :B1, 20
# c1 = spreadsheet.add_cell :C1, '= (A1 + B1) * 2'
# d1 = spreadsheet.add_cell :D1, '= C1 - 1'
#
# puts c1.references.map(&:ref).inspect
# puts a1.observers.map(&:ref).inspect
# puts b1.observers.map(&:ref).inspect
# puts c1.content
# puts c1.eval
# puts d1.eval
#
# c1.content = '= (A2 + B1) ** 2'
# puts a1.observers.map(&:ref).inspect
# puts b1.observers.map(&:ref).inspect
# b2 = spreadsheet.find_or_create_cell_by_ref(:a2)
# puts b2.observers.map(&:ref).inspect
# puts c1.content
# puts c1.eval
#
# a1.content = 1
# puts c1.eval
#
# b1.content = 2
# puts c1.eval
#
# b1.content = -2
# puts c1.eval
# puts d1.eval

a1 = spreadsheet.add_cell :A1, 10
b1 = spreadsheet.add_cell :B1, '= A1 + 1'
c1 = spreadsheet.add_cell :C1, '= B1 + 2'
d1 = spreadsheet.add_cell :D1, '= C1 + 3'
e1 = spreadsheet.add_cell :E1, '       = D1 + 4        '

puts e1.eval

a1.content = 20
puts e1.eval

c1.content = 100
puts e1.eval

c1.content = '= B1 + 200'
puts e1.eval

a1.content = '= E1'
