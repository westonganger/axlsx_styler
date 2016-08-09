require 'axlsx'

require 'axlsx_styler/version'
require 'axlsx_styler/axlsx_workbook'
require 'axlsx_styler/axlsx_worksheet'
require 'axlsx_styler/axlsx_cell'

Axlsx::Workbook.send :include, AxlsxStyler::Axlsx::Workbook
Axlsx::Worksheet.send :include, AxlsxStyler::Axlsx::Worksheet
Axlsx::Cell.send :include, AxlsxStyler::Axlsx::Cell

module Axlsx
  class Package
    # Patches the original Axlsx::Package#serialize method so that styles are
    # applied when the workbook is saved
    original_serialize = instance_method(:serialize)
    define_method :serialize do |*args|
      workbook.apply_styles if !workbook.styles_applied
      original_serialize.bind(self).(*args)
    end

    # Patches the original Axlsx::Package#to_stream method so that styles are
    # applied when the workbook is converted to StringIO
    original_to_stream = instance_method(:to_stream)
    define_method :to_stream do |*args|
      workbook.apply_styles if !workbook.styles_applied
      original_to_stream.bind(self).(*args)
    end
  end

  class Styles
    # Patches the original Axlsx::Styles#add_style_method so that we can combine 
    # the existing cell style with the new one by adding the optional 'style_index' argument
    def add_style(options={}, style_index=nil)
      options[:type] ||= :xf # Default to :xf
      raise ArgumentError, "Type must be one of [:xf, :dxf]" unless [:xf, :dxf].include?(options[:type])

      fill = parse_fill_options(options)
      font = parse_font_options(options)
      numFmt = parse_num_fmt_options(options)
      border = parse_border_options(options)
      alignment = parse_alignment_options(options)
      protection = parse_protection_options(options)

      if options[:type] == :dxf
        style = Dxf.new(fill: fill, font: font, numFmt: numFmt, border: border, alignment: alignment, protection: protection)
      else
        unless style_index.nil?
          style_object = cellXfs[style_index]

          if font.nil?
            font = style_object.fontId
          else
            font = parse_font_options(options, style_object.fontId)
          end

          fill ||= style_object.fillId
          numFmt ||= style_object.numFmtId
          border ||= style_object.borderId
          alignment ||= style_object.alignment
          protection ||= style_object.protection
        end

        style = Xf.new(fillId: (fill || 0), fontId: (font || 0), numFmtId: (numFmt || 0), borderId: (border || 0), alignment: alignment, protection: protection, applyFill: !fill.nil?, applyFont: !font.nil?, applyNumberFormat: !numFmt.nil?, applyBorder: !border.nil?, applyAlignment: !alignment.nil?, applyProtection: !protection.nil?)
      end

      if options[:type] == :xf
        cellXfs << style
      else
        dxfs << style
      end
    end

    # allow to pass in fontId instead of always defaulting first, also apply name and color before merging options
    def parse_font_options(options={}, fontId=0)
      return if (options.keys & [:fg_color, :sz, :b, :i, :u, :strike, :outline, :shadow, :charset, :family, :font_name]).empty?

      options[:name] = options[:font_name] unless options[:font_name].nil?
      options[:color] = Color.new(rgb: options[:fg_color]) unless options[:fg_color].nil?

      options = fonts[fontId].instance_values.merge(options)
      font = Font.new(options)

      if options[:type] == :dxf
        font
      else
        fonts << font
      end
    end
  end
end
