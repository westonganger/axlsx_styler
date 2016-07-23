module AxlsxStyler
  module Axlsx
    module Styles
      def add_style(options={}, style_object=nil)
        options[:type] ||= :xf # Default to :xf
        raise ArgumentError, "Type must be one of [:xf, :dxf]" unless [:xf, :dxf].include?(options[:type])

        fill = parse_fill_options(options)
        font = parse_font_options(options)
        numFmt = parse_num_fmt_options(options)
        border = parse_border_options(options)
        alignment = parse_alignment_options(options)
        protection = parse_protection_options(options)

        if style_object
          fill ||= style_object.fillId
          font ||= style_object.fontId
          numFmt ||= style_object.numFmtId
          border ||= style_object.borderId
          alignment ||= style_object.alignment
          protection ||= style_object.protection
        end

        if options[:type] == :dxf
          style = Dxf.new(fill: fill, font: font, numFmt: numFmt, border: border, alignment: alignment, protection: protection)
        else
          style = Xf.new(fillId: (fill || 0), fontId: (font || 0), numFmtId: (numFmt || 0), borderId: (border || 0), alignment: alignment, protection: protection, applyFill: !fill.nil?, applyFont: !font.nil?, applyNumberFormat: !numFmt.nil?, applyBorder: !border.nil?, applyAlignment: !alignment.nil?, applyProtection: !protection.nil?)
        end

        if options[:type] == :xf
          cellXfs << style
        else
          dxfs << style
        end
      end
    end
  end
end
