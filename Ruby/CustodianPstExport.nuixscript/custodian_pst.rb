# Menu Title: Custodian PST Export
# Needs Case: true
# Needs Selected Items: true
# @version 2.1.2

naming = 'item_name'
# naming = 'item_name_with_path'

load File.join(__dir__, 'nx_progress.rb') # v1.0.0
load File.join(__dir__, 'summary_reporter.rb') # v1.0.0

require 'fileutils'

# Class for exporting items by custodian.
# * +@export_dir+ is the export directory
# * +@items_dir+ is the export items directory
# * +@reports_dir+ is the reports directory
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
    @items_dir = 'Items'
    @reports_dir = 'Reports'
    @naming = naming
    ProgressDialog.forBlock do |progress_dialog|
      super(progress_dialog, 'Custodian Export')
      @top_items = tops(items)
      @total = @top_items.size
      run(export_dir) unless custodian_missing
    end
  end

  protected

  # Creates top-level-MD5-digests.txt by combining per-custodian files.
  def append_digests
    @@dialog.setSubStatusAndLogIt('Generating top-level-MD5-digests.txt')
    files = Dir.glob(report_path(File.join('*', 'top-level-MD5-digests.txt')))
    @@dialog.setSubProgress(0, files.size)
    files.each_with_index do |file, i|
      @@dialog.setSubProgress(i)
      @@dialog.logMessage("Adding #{file}")
      o = File.new(File.join(@export_dir, 'top-level-MD5-digests.txt'), 'a')
      o.write(File.read(file))
      o.close
    end
  end

  # Creates BatchExporter with product and numbering options.
  #
  # @return [BatchExporter]
  def create_exporter(name)
    export_options = { naming: @naming, mailFormat: 'pst', path: name }
    exporter = $utilities.create_batch_exporter(export_path(nil))
    exporter.add_product('native', export_options)
    exporter.set_numbering_options(createProductionSet: false) if $utilities.get_license.has_feature('PRODUCTION_SET')
    exporter
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
  #
  # @return [true] once all items have been exported
  # @return [nil] if abort was requested
  def export
    @@dialog.setMainStatusAndLogIt("Exporting to #{@export_dir}")
    custodians = $current_case.get_all_custodians
    @@dialog.setSubProgress(0, custodians.size)
    custodians.each_with_index do |c, index|
      @@dialog.setSubProgress(index)
      return true if export_custodian(c)
      return nil if @@dialog.abortWasRequested
    end
  end

  # Exports items for custodian.
  #
  # @param name [String] custodian of the items
  # @return [true, false] if all items have been exported
  def export_custodian(name)
    items = intersect_top("custodian:\"#{name}\"")
    return nil if items.empty?

    i = items.size
    @@dialog.setSubStatusAndLogIt("Exporting custodian: #{name} (#{i} items)")
    create_exporter(name).export_items(items)
    export_rename(name)
    progress_check(i)
  end

  # Returns a path inside is export items directory.
  #
  # @param name [String, nil] the string to add, or nil to return the base path
  # @return [String] the path for the export items
  def export_path(name)
    return File.join(@export_dir, @items_dir) if name.nil?

    File.join(@export_dir, @items_dir, name)
  end

  # Renames files from custodian export.
  #  Moves files if they aren't a PST.
  #  Moves/renames PST (and deletes folder if it's now empty).
  #
  # @param custodian [String]
  def export_rename(custodian)
    @@dialog.logMessage('Renaming files')
    dir = report_path(custodian)
    FileUtils.mkdir_p(dir)
    Dir.glob(export_path('*')).each do |p|
      if File.file?(p)
        rename_file(p, File.join(dir, File.basename(p))) unless p.end_with?('.pst')
      elsif @naming == 'item_name'
        rename_psts(p, custodian)
      end
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

  # Renames file.
  #
  # @param old [String] initial file
  # @param new [String] destination
  def rename_file(old, new)
    @@dialog.logMessage("Creating #{new}")
    File.rename(old, new)
  end

  # Renames exported PSTs to the custodian name.
  #
  # @param dir [String] the path containing export PST
  # @param custodian [String]
  def rename_psts(dir, custodian)
    # Handle multiple PSTs
    Dir.glob(File.join(dir, 'Export*.pst')).each do |pst|
      rename_file(pst, export_path(File.basename(pst).sub('Export', custodian)))
    end
    # Remove directory if now empty
    FileUtils.remove_dir(dir) if Dir.glob(File.join(dir, '*')).empty?
  end

  # Returns a path inside is reports directory.
  #
  # @param name [String, nil] the string to add, or nil to return the base path
  # @return [String] the report path
  def report_path(name)
    return File.join(@export_dir, @reports_dir) if name.nil?

    File.join(@export_dir, @reports_dir, name)
  end

  # Exports items per-custodian, then moves and summarizes reports.
  #
  # @param export_dir [String] directory for export
  def run(export_dir)
    @@dialog.setMainProgress(0, @total + 4)
    start = Time.now
    export
    SummaryReporter.new(start, export_dir, report_path(nil), 'Custodian').write
    advance_main
    append_digests
    close_nx unless @@dialog.abortWasRequested
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
