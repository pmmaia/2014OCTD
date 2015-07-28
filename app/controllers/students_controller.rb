class StudentsController < ApplicationController
  before_filter :load_students
  before_filter :load_workbook, except: [:index, :generate]
  def index
  end

  def generate
    %x[rake generate:data] 
  end

  def export
    case params[:type]
    when "Development"
      export_all_together_xlsx
    end
  end

  def export_basic_xlsx
    @wb.add_worksheet(name: "Basic") do |sheet|
      sheet.add_row get_header 
      @students.each do |st|
        sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
      end
    end
    @p.serialize("#{Rails.root}/tmp/basic.xlsx")
    send_file("#{Rails.root}/tmp/basic.xlsx", filename: "Basic.xlsx", type: "application/xlsx")
  end

  def export_row_col_xlsx
    @wb.add_worksheet(name: "Row&Col") do |sheet|
      sheet.add_row get_header 
      @students.each do |st|
        sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
      end
      sheet.col_style 4, @center, row_offset: 1
      sheet.row_style 0, @header, col_offset: 1
    end
    @p.serialize("#{Rails.root}/tmp/row_col.xlsx")
    send_file("#{Rails.root}/tmp/row_col.xlsx", filename: "Row_Col.xlsx", type: "application/xlsx")
  end

  def export_custom_xlsx
    @p.use_autowidth = false
    @wb.add_worksheet(name: "Custom") do |sheet|
      sheet.add_row get_header, style: @header
      @students.each do |st|
        if st.fname.length >= 21
          sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @data, height: 25 
        else
          sheet.add_row [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark], style: @data 
        end
      end
      sheet.column_widths 20, 20, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/custom.xlsx")
    send_file("#{Rails.root}/tmp/custom.xlsx", filename: "Custom.xlsx", type: "application/xlsx")
  end

  def export_all_together_xlsx
    @wb.add_worksheet(name: "All") do |sheet|
      sheet.add_row ["Draft Offsite Waste Disposal Classification Indicator", "", "", "", "", ""], style: @heading, height: 30
      sheet.merge_cells("A1:F1")
      sheet.add_row get_header, style: @header
      @students.each do |st|
        if st.fname.length >= 21
          if st.remark == "California Hazardous"
            sheet.add_row [st.fname, st.lname, st.grade, st.marks, st.percentage, st.remark], style: @style_pass, height: 25
          else
            sheet.add_row [st.fname, st.lname, st.grade, st.marks, st.percentage, st.remark], style: @style_fail, height: 25
          end
        else
          if st.remark == "California Hazardous"
            sheet.add_row [st.fname, st.lname, st.grade, st.marks, st.percentage, st.remark], style: @style_pass
          else
            sheet.add_row [st.fname, st.lname, st.grade, st.marks, st.percentage, st.remark], style: @style_fail
          end
        end
      end
        sheet.add_row ["", "", "Min", "=MIN(D3:D102)", "=MIN(E3:E102)", ""], style: @style_fail
        sheet.add_row ["", "", "Max", "=MAX(D3:D102)", "=MAX(E3:E102)", ""], style: @style_fail
        sheet.add_row ["", "", "Average", "=AVERAGE(D3:D102)", "", ""], style: @style_fail
      sheet.column_widths 20, 20, nil, nil, nil, 20
    end
    @p.serialize("#{Rails.root}/tmp/all.xlsx")
    send_file("#{Rails.root}/tmp/all.xlsx", filename: "All.xlsx", type: "application/xlsx")
  end

  def export_merge_xlsx 
    @wb.add_worksheet(name: "All") do |sheet|
      sheet.add_row ["Student Result Detail", "", "", "", "", ""], style: @heading, height: 30
      sheet.merge_cells("A1:F1")
      sheet.add_row get_header, style: @header
      @students_with_a = Student.where(grade: "A") 
      @students_with_b = Student.where(grade: "B") 
      @students_with_c = Student.where(grade: "C")
      @students_with_f = Student.where(grade: "")
      @students_with_a.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_pass, height: 25  
        else
          sheet.add_row data_array, style: @style_pass 
        end
      end
      a = @students_with_a.length
      sheet.add_row ["", "Students With Grade A", "=AVERAGE(C3:C#{a+2})", "=AVERAGE(D3:D#{a+2})", "Total", a], style: @total

      @students_with_b.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_pass, height: 25  
        else
          sheet.add_row data_array, style: @style_pass 
        end
      end
      b = @students_with_b.length
      sheet.add_row ["", "Students With Grade B", "=AVERAGE(C#{a+4}:C#{a+b+3})", "=AVERAGE(D#{a+4}:D#{a+b+3})", "Total", b], style: @total

      @students_with_c.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_pass, height: 25  
        else
          sheet.add_row data_array, style: @style_pass 
        end
      end
      c = @students_with_c.length
      sheet.add_row ["", "Students With Grade C", "=AVERAGE(C#{a+b+4}:C#{a+b+c+4})", "=AVERAGE(D#{a+b+4}:D#{a+b+c+4})", "Total", c], style: @total

      @students_with_f.each do |st|
        data_array = [st.fname, st.lname, st.marks, st.percentage, st.grade, st.remark]
        if st.fname.length >= 21
          sheet.add_row data_array, style: @style_fail, height: 25  
        else
          sheet.add_row data_array, style: @style_fail 
        end
      end
      f = @students_with_f.length
      sheet.add_row ["", "Failed Students", "=AVERAGE(C#{a+b+c+4}:C#{a+b+c+f+4})", "=AVERAGE(D#{a+b+c+4}:D#{a+b+c+f+4})", "Total", f], style: @total

      sheet.column_widths 20, 20, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/Merge.xlsx")
    send_file("#{Rails.root}/tmp/Merge.xlsx", filename: "Merge.xlsx", type: "application/xlsx")
  end

  def export_image_xlsx
    @wb.add_worksheet(name: "Image") do |sheet|
      sheet.add_row ["", "Yehh !! Results", "", "", "", ""], style: @heading, height: 30
      img = File.expand_path(Rails.root+'app/assets/images/result.png')
      sheet.add_image(:image_src => img, :hyperlink=>"http://rubyInsense.heroku.com") do |image|
        image.width=400
        image.height=300
        image.hyperlink.tooltip = "Labeled Link"
        image.start_at 1, 1
      end
    end
    @wb.add_worksheet(name: "Data Type") do |sheet|
      sheet.add_row ["Date", "Time", "String", "Boolean", "Float", "Integer"]
      sheet.add_row [Date.today, Time.now, "value", true, 0.1, 1], :style => [@date_format, @time_format]
      sheet.column_widths 10, 10, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/Image.xlsx")
    send_file("#{Rails.root}/tmp/Image.xlsx", filename: "Image.xlsx", type: "application/xlsx")
  end

  def export_hyperlink_xlsx
    @wb.add_worksheet(:name => 'Hyperlinks') do |sheet|
      sheet.add_row ['rubyInsense']
      sheet.add_hyperlink :location => 'http://rubyInsense.heroku.com', :ref => sheet.rows.first.cells.first
      sheet.add_hyperlink :location => "'Next Sheet'!A1", :ref => 'A2', :target => :sheet
      sheet.add_row ['Go to next sheet']
    end
    @wb.add_worksheet(:name => 'Next Sheet') do |sheet|
      sheet.add_row ['hello!']
    end
    @p.serialize("#{Rails.root}/tmp/links.xlsx")
    send_file("#{Rails.root}/tmp/links.xlsx", filename: "Links.xlsx", type: "application/xlsx")
  end

  def export_bar_chart_axlsx
    @a = Student.where(grade: "A").count
    @b = Student.where(grade: "B").count
    @c = Student.where(grade: "C").count
    @fail = Student.where(remark: "FAIL").count
    @wb.add_worksheet(name: "Bar Chart") do |sheet|
      sheet.add_row ["", "Result Analysis", "", "", "", ""], style: @heading, height: 30
      sheet.add_row ["Grade A", "Grade B", "Grade C", "FAIL"]
      sheet.add_row [@a, @b, @c, @fail]
      sheet.add_chart(Axlsx::Bar3DChart, :start_at => "A4", :end_at => "H19", :bar_dir => :col) do |chart|
        chart.add_series :data => sheet["A3:D3"], :labels => sheet["A2:D2"], :title => sheet["B1"], colors: ["00FF00", "0066CC", "F0", "FF0000"]
      end
      sheet.column_widths 10, 10, nil, nil, nil, nil
    end
    @p.serialize("#{Rails.root}/tmp/bar.xlsx")
    send_file("#{Rails.root}/tmp/bar.xlsx", filename: "Bar.xlsx", type: "application/xlsx")
  end

  def export_line_chart_axlsx
    @wb.add_worksheet(:name => "Line Chart") do |sheet|
      sheet.add_row ["First", 1, 5, 7, 9]
      sheet.add_row ["Second", 5, 2, 14, 9]
      sheet.add_chart(Axlsx::LineChart, :title => "Line Chart") do |chart|
        chart.start_at 0, 2
        chart.end_at 10, 15
        chart.add_series :data => sheet["B1:E1"], :title => sheet["A1"], :color => "0000FF"
        chart.add_series :data => sheet["B2:E2"], :title => sheet["A2"], :color => "FF0000"
        chart.catAxis.title = 'Y Axis'
        chart.valAxis.title = 'X Axis'
      end
    end
    @p.serialize("#{Rails.root}/tmp/line.xlsx")
    send_file("#{Rails.root}/tmp/line.xlsx", filename: "line.xlsx", type: "application/xlsx")
  end

  def export_pie_chart_axlsx
    @wb.add_worksheet(:name => "Pie Chart") do |sheet|
      sheet.add_row ["", "Result Analysis"], style: @heading
      sheet.add_row ["Grade", "Percentage"], style: @header
      @a = Student.where(grade: "A").count
      @b = Student.where(grade: "B").count
      @c = Student.where(grade: "C").count
      sheet.add_row ["A", @a]
      sheet.add_row ["B", @b]
      sheet.add_row ["C", @c]
      sheet.add_chart(Axlsx::Pie3DChart, :start_at => [0,6], :end_at => [6, 20], :title => "Pie Chart") do |chart|
        chart.add_series :data => sheet["B3:B5"], :labels => sheet["A3:A5"],  :colors => ['FF0000', '00FF00', '0000FF']
        chart.d_lbls.d_lbl_pos = :bestFit
        chart.d_lbls.show_percent = :true
      end
    end
    @p.serialize("#{Rails.root}/tmp/pie.xlsx")
    send_file("#{Rails.root}/tmp/pie.xlsx", filename: "pie.xlsx", type: "application/xlsx")
  end

  private
  def load_students
    @students = Student.all
  end

  def load_workbook
    @p = Axlsx::Package.new
    @wb = @p.workbook
    load_styles
  end

  def load_styles
    @wb.styles do |s| 
      @item_style = s.add_style :b => false, :sz => 9,  :font_name => 'Century Gothic', :alignment => { :horizontal => :left, :vertical => :center, :wrap_text => true}
      @heading = s.add_style alignment: {horizontal: :center}, b: true, sz: 18, bg_color: "0066CC", fg_color: "FF", :font_name => "Century Gothic"
      @header = s.add_style alignment: {horizontal: :left}, b: true, sz: 10, bg_color: "4F628E", :font_name => "Comic Sans Ms"
      @data_str_red= s.add_style alignment: {wrap_text: true}, bg_color: "D4C26A"
      @data_str_green= s.add_style alignment: {wrap_text: true}, bg_color: "99A637"
      @data_num_red= s.add_style alignment: {wrap_text: true}, b: true, i: true, alignment: { horizontal: :left, vertical: :center }, bg_color: "D4C26A"
      @data_num_green= s.add_style alignment: {wrap_text: true}, b: true, i: true, alignment: { horizontal: :left, vertical: :center }, bg_color: "99A637"
      @center = s.add_style alignment: {horizontal: :center}, fg_color: "0000FF" 
      @green = s.add_style alignment: {horizontal: :left}, fg_color: "000000", bg_color: "99A637" 
      @red = s.add_style alignment: {horizontal: :left}, fg_color: "FF0000", bg_color: "D4C26A", b: true
      @total = [@data, @header, @header, @header, @header, @header]
      @style_pass = [@data_str_red, @data_str_red, @data_str_red, @data_num_red, @data_num_red, @red]
      @style_fail = [@data_str_green, @data_str_green, @data_str_green, @data_num_green, @data_num_green, @green]
      @date_format = s.add_style :format_code => 'YYYY-MM-DD'
      @time_format = s.add_style :format_code => 'hh:mm:ss'
    end
  end

  def get_header
    ["Laboratory ID", "Sample ID", "Scheme", "Pb (mg/Kg)", "Wet Pb (mg/L)", "Classification"]
  end
end
