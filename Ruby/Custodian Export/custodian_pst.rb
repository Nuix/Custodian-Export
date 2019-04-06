# Menu Title: Custodian PST Export
# Needs Case: true
# Needs Selected Items: true
# @version 2.0.0

naming = 'item_name'
# naming = 'item_name_with_path'

load File.join(__dir__, 'nx_progress.rb_') # v1.0.0
load File.join(__dir__, 'summary_reporter.rb_') # v1.0.0

require 'fileutils'

# Class for exporting items by custodian.
# * +@export_dir+ is the export directory
# * +@path[:exports]+ is the exports path
# * +@path[:reports]+ is the reports path
# * +@naming+ is the naming property for the export product
# * +@top_items+ are the top level items that will be exported
# * +@total+ is @top_items.size
class CustodianExport < NxProgress
  # Exports items by custodian and creates summary report.
  #
  # @param export_dir [String] path for exports
  # @param items [Collection<Item>] items to export
  # @param naming [String] naming scheme for export product
  def initialize(export_dir, items, naming)
    @export_dir = export_dir
    @path = paths(export_dir)
    @naming = naming
    ProgressDialog.forBlock do |progress_dialog|
      super(progress_dialog, 'Custodian Export')
      @top_items = tops(items)
      @total = @top_items.size
      run(export_dir) unless custodian_missing
      close_nx
    end
  end

  protected

  # Creates top-level-MD5-digests.txt by combining per-custodian files.
  def append_digests
    name = 'top-level-MD5-digests.txt'
    @@dialog.setSubStatusAndLogIt("Generating #{name}")
    files = Dir.glob(File.join(@path[:reports], '*', name))
    @@dialog.setSubProgress(0, files.size)
    files.each_with_index do |file, i|
      @@dialog.setSubProgress(i)
      @@dialog.logMessage("Adding #{file}")
      o = File.new(File.join(@export_dir, name), 'a')
      o.write(File.read(file))
      o.close
    end
  end

  # Exports items for custodian.
  #
  # @param name [String] custodian of the items
  # @param items [Collection<Item>] items to export
  # @return [true, false] if all items have been exported
  def batch_export(name, items)
    i = items.size
    @@dialog.setSubStatusAndLogIt("Exporting custodian: #{name} (#{i} items)")
    export_options = { naming: @naming, mailFormat: 'pst' }
    path = File.join(@path[:exports], name)
    exporter = $utilities.create_batch_exporter(path)
    exporter.add_product('native', export_options)
    exporter.set_numbering_options(createProductionSet: false)
    exporter.export_items(items)
    progress_check(i)
  end

  # Checks if +@top_items+ all have a custodian.
  # If there are items without custodian, logs GUIDs.
  #
  # @return [true, false] true if all +@export_top+ items have a custodian
  def custodian_missing
    @@dialog.logMessage("#{@top_items.size} top-level items to export")
    missing_custodian = intersect_top('has-custodian:0')
    return false if missing_custodian.empty?

    @@dialog.setMainStatusAndLogIt('ERROR - Items missing custodian')
    missing_custodian.each { |i| puts @@dialog.logMessage(i.get_guid) }
    true
  end

  # Exports items.
  def export
    @@dialog.setMainStatusAndLogIt("Exporting to #{@path[:exports]}")
    custodians = $current_case.get_all_custodians
    @@dialog.setSubProgress(0, custodians.size)
    custodians.each_with_index do |c, index|
      @@dialog.setSubProgress(index)
      items = intersect_top("custodian:\"#{c}\"")
      next if items.empty?
      break if @@dialog.abortWasRequested || batch_export(c, items)
    end
  end

  # Computes the intersection of items and +@top_items+.
  # Returns the result as a new set.
  #
  # @param query [String] search query
  # @return [Set<Item>] the items also contained in +@top_items+
  def intersect_top(query)
    items = $current_case.search_unsorted(query)
    $utilities.get_item_utility.intersection(@top_items, items)
  end

  # Moves files to reports directory.
  # Creates custodian directory and logs messages.
  #
  # @param dir [String] directory containg files to move
  def move(dir)
    @@dialog.logMessage("Moving files from #{dir}")
    # Make per-custodian directory
    FileUtils.mkdir_p(File.join(@path[:reports], File.basename(dir)))
    ['top-level-MD5-digests.txt', 'summary-report.xml', 'summary-report.txt'].each do |r|
      f = File.join(dir, r)
      n = f.sub(@path[:exports], @path[:reports])
      @@dialog.logMessage("Moving #{r} to #{n}")
      File.rename(f, n)
    end
  end

  # Moves reports.
  def move_reports
    @@dialog.setMainStatusAndLogIt('Preparing summary report')
    @@dialog.setSubStatusAndLogIt("Moving reports to #{@path[:reports]}")
    to_move = Dir.glob(File.join(@path[:exports], '*', File::SEPARATOR))
    @@dialog.setSubProgress(0, to_move.size)
    to_move.each_with_index do |dir, i|
      move(dir)
      @@dialog.setSubProgress(i)

      break if @@dialog.abortWasRequested
    end
    advance_main
  end

  def paths(export_dir)
    { exports: File.join(export_dir, 'Export'),
      reports: File.join(export_dir, 'Reports') }
  end

  # Advances progress, outputs, and returns true when completed.
  #
  # @param count [Integer] number of items to advance
  # @return [true, false] true once progress is complete
  def progress_check(count)
    @@progress += count
    @@dialog.setMainProgress(@@progress)
    @@dialog.logMessage("Exported #{@@progress} of #{@total} items")
    @@progress == @total
  end

  # Exports items per-custodian, then moves and summarizes reports.
  #
  # @param export_dir [String] directory for export
  def run(export_dir)
    @path.each { |k, v| @@dialog.logMessage("Writing #{k} to #{v}") }
    @@dialog.setMainProgress(0, @total + 4)
    start = Time.now
    return false if export == false || @@dialog.abortWasRequested

    move_reports
    return false if @@dialog.abortWasRequested

    SummaryReporter.new(start, export_dir, @path[:reports], 'Custodian').write
    advance_main
    append_digests
  end

  # Finds the top level items from selected items.
  #
  # @param items [Collection<Item>] items to export
  # @return [Collection<Item> top-level items
  def tops(items)
    @@dialog.setMainStatusAndLogIt('Getting items to export')
    @@dialog.setSubStatusAndLogIt("#{items.size} selected items")
    @@dialog.setSubProgress(0, 2)
    $utilities.get_item_utility.find_top_level_items(items, true)
  end
end

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
  CustodianExport.new(dir, $current_selected_items, naming) unless dir.nil?
end
