# Menu Title: Custodian Export
# Needs Case: true
# Needs Selected Items: true

require File.join(__dir__, 'custodian_export.rb')

# Prompts user to choose a directory.
#
# @return [String, nil] absolute path or nil if canceled
def choose_export_dir
  java_import javax.swing.JFileChooser
  c = JFileChooser.new
  c.setFileSelectionMode(JFileChooser::DIRECTORIES_ONLY)
  c.setDialogTitle('Select Export Directory')
  return nil unless c.showOpenDialog(nil) == JFileChooser::APPROVE_OPTION

  c.getSelectedFile.getAbsolutePath
end

begin
  dir = choose_export_dir
  CustodianExport.new(dir, $current_selected_items) unless dir.nil?
end
