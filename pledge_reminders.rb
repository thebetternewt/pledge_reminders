require 'csv'
require 'spreadsheet'

#########################################################################
# Helper Functions
#########################################################################

def timestamp
  DateTime.now.strftime('%y%m%dT%H%M%S%z')
end

def to_sheet(sheet, records)
  records.each_with_index do |r, i|
    sheet.update_row (i + 1),
    r['MSU-ID'],
    r['NAME'],
    r['Reminder Type'],
    r['PLEDGE_NUMBER'],
    r['PLEDGE_DATE'],
    r['PLEDGE_TYPE'],
    r['AMOUNT_PLEDGED'],
    r['PLEDGE_AMOUNT_PAID'],
    r['PLEDGE_BALANCE'],
    r['PLEDGE_AMT_DUE'],
    r['COLL_CODE'],r['AREA'],
    r['DESG1'],
    r['DESG2'],
    r['DESG3'],
    r['DESG4'],
    r['DESG5'],
    r['DESG6'],
    r['DESG7'],
    r['DESG8'],
    r['DESG9'],
    r['DESG10'],
    r['STREET_LINE1'],
    r['STREET_LINE2'],
    r['STREET_LINE3'],
    r['CITY'],
    r['STATE'],
    r['ZIP'],
    r['ALUM_EMAIL'],
    r['SOLC_TYPE'],
    r['SOLC_ORG']
  end
end

#########################################################################
# Setup
#########################################################################

# Cleanup old .xls files in directory.
FileUtils.rm Dir.glob('*.xls')

# Create college codes and groups dictionary (hash).
college_codes = { '00' => { name: 'General University',
                            members: %w[00 10 12 15 99] },
                  '02' => { name: 'AG', members: %w[02 14] },
                  '03' => { name: 'CAAD', members: %w[03] },
                  '04' => { name: 'A&S', members: %w[04] },
                  '05' => { name: 'COB', members: %w[01 05] },
                  '06' => { name: 'ED', members: %w[06] },
                  '07' => { name: 'ENG', members: %w[07] },
                  '08' => { name: 'FR', members: %w[08] },
                  '09' => { name: 'VM', members: %w[09] },
                  '11' => { name: 'Grad School', members: %w[11 16] },
                  '13' => { name: 'Meridian', members: %w[13] } }

#########################################################################
# Read and clean up raw data from pledge reminders file.
#########################################################################

# Assign first argument as input_file path.
input_file = ARGV.first

# headers = CSV.read(input_file, headers: true)
reminders = CSV.read(input_file,
                     headers: true,
                     skip_lines: /^[,\s]+$|rows selected/) # Skip last two rows.

# Clean up trailing spaces in values.
reminders.each do |row|
  row.values_at.each do |v|
    v.to_s.rstrip!
  end
end

# Find college codes present in current report.
present_codes = reminders.values_at('COLL_CODE').uniq.flatten
present_codes = present_codes.map(&:to_i).sort!.map(&:to_s)
p present_codes

# Find groups present in current report.
present_groups = []
college_codes.each do |group_code, data|
  data.each do |_name, members|
    present_codes.each do |code|
      present_groups << group_code if members.include?(code.rjust(2, '0'))
      present_groups.uniq!
    end
  end
end
p present_groups

#########################################################################
# Generate Excel workbook.
#########################################################################

book = Spreadsheet::Workbook.new

header_format = Spreadsheet::Format.new color: :white,
                                        weight: :bold,
                                        pattern_fg_color:
                                          :xls_color_29, # maroon
                                        pattern: 1

# Create sheets from college codes.
present_groups.each do |group_code|
  sheet_name = "#{group_code}-#{college_codes[group_code][:name]}"
  new_sheet = book.create_worksheet name: sheet_name

  # Select reminders with college code.
  selected_reminders = reminders.select do |r|
    college_codes[group_code][:members].include?(r['COLL_CODE'].rjust(2, '0'))
  end
  selected_reminders.sort_by! { |r| r['AREA'] }

  # Write sheet headers.
  reminders.headers.each { |header| new_sheet.row(0).push header }
  new_sheet.row(0).default_format = header_format

  # Write reminders to sheet.
  to_sheet(new_sheet, selected_reminders)
  new_sheet.columns.autofit
end

#########################################################################
# Catch reminders not included in known groups.
#########################################################################

new_sheet = book.create_worksheet name: 'OTHER'

# Select reminders with unknown college code.
known_codes = college_codes.values.collect { |code| code[:members] }
known_codes.flatten!.uniq!
selected_reminders = reminders.reject do |r|
  known_codes.include?(r['COLL_CODE'].rjust(2, '0'))
end
# selected_reminders = reminders.select { |r|  }
selected_reminders.sort_by! { |r| r['AREA'] }

# Write sheet headers.
reminders.headers.each { |header| new_sheet.row(0).push header }
new_sheet.row(0).default_format = header_format

# Write reminders to sheet.
to_sheet(new_sheet, selected_reminders)

##########################################################################
# Generate Excel spreadsheet.
##########################################################################

wb_name = "devoff_pldg_reminders_#{timestamp}.xls"
book.write "./#{wb_name}"

# Open file upon completion.
if /cygwin|mswin|mingw|bccwin|wince|emx/ =~ RUBY_PLATFORM # Check if Windows OS
  system %(cmd /c "start #{wb_name}")
else system %(open "#{wb_name}") # Assume Mac OS/Linux
end
