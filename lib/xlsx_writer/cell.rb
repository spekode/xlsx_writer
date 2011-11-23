require 'fast_xs'

module XlsxWriter
  class Cell
    class << self
      # TODO make a class for this
      def excel_type(calculated_type)
        case calculated_type
        when :String
          :inlineStr
        when :Number, :Date, :Currency
          :n
        when :Boolean
          :b
        else
          raise ::ArgumentError, "Unknown cell type #{k}"
        end
      end
      
      # TODO make a class for this
      def excel_style_number(calculated_type)
        case calculated_type
        when :String
          0
        when :Number
          0
        when :Currency
          1
        when :Date
          2
        when :Boolean
          0 # todo
        else
          raise ::ArgumentError, "Unknown cell type #{k}"
        end
      end
      
      def excel_column_letter(i)
        result = []
        while i >= 26 do
          result << ABC[i % 26]
          i /= 26
        end
        result << ABC[result.empty? ? i : i - 1]
        result.reverse.join
      end
            
      def excel_string(value)
        value.to_s.fast_xs
      end
      
      def excel_number(value)
        str = value.to_s.dup
        unless str =~ /\A[0-9\.]*\z/
          raise ::ArgumentError, %{Bad value "#{value}" Only numbers and dots (.) allowed in number fields}
        end
        str.fast_xs
      end
      
      alias :excel_currency :excel_number
      
      # doesn't necessarily work for times yet
      JAN_1_1900 = ::Time.parse('1900-01-01')
      def excel_date(value)
        if value.is_a?(::String)
          ((::Time.parse(str) - JAN_1_1900) / 86_400).round
        elsif value.respond_to?(:to_date)
          (value.to_date - JAN_1_1900.to_date).to_i
        end
      end
      
      def excel_boolean(value)
        value ? 1 : 0
      end
    end
    
    ABC = ('A'..'Z').to_a
    
    attr_reader :row
    attr_reader :data
    
    def initialize(row, data)
      @row = row
      @data = data.is_a?(::Hash) ? data.symbolize_keys : data
    end
    
    def unstyled?
      !styled?
    end
    
    def styled?
      data.is_a?(::Hash)
    end
    
    def to_xml
      if value.blank?
        %{<c r="#{excel_column_letter}#{row.ndx}" s="0" t="inlineStr" />}
      elsif excel_type == :inlineStr
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="#{excel_type}"><is><t>#{excel_value}</t></is></c>}
      else
        %{<c r="#{excel_column_letter}#{row.ndx}" s="#{excel_style_number}" t="#{excel_type}"><v>#{excel_value}</v></c>}
      end
    end
    
    # 0 -> A (zero based!)
    def excel_column_letter
      Cell.excel_column_letter row.cells.index(self)
    end
    
    # detect dates here, even if we're not styled
    def excel_type
      Cell.excel_type calculated_type
    end
    
    def excel_style_number
      Cell.excel_style_number calculated_type
    end

    def calculated_type
      if styled?
        data[:type]
      elsif value.is_a?(::Date)
        :Date
      elsif value.is_a?(::Numeric) or value.to_s =~ /\A[0-9\.,]+\z/
        :Number
      else
        :String
      end
    end
    
    def value
      styled? ? data[:value] : data
    end
    
    def excel_value
      Cell.send "excel_#{calculated_type.to_s.underscore}", value
    end
  end
end
