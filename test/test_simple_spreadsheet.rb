require_relative '../simple_spreadsheet'
require 'contest'
require 'turn/autorun'

class TestSpreadSheet < Test::Unit::TestCase
  context 'SpreadSheet' do
    setup do
      @spreadsheet = SpreadSheet.new
    end

    context 'Cell' do
      setup do
        @developer = ["Cassiano D'Andrea <cassiano.dandrea@tagview.com.br>", Time.now.utc, '-0300']
      end

      test '#add_cell' do
        a1 = spreadsheet.add_cell :A1, 10

        ...
      end
    end
  end
end
