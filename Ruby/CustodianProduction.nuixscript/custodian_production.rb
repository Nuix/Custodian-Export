# Menu Title: Custodian Production
# Needs Case: true
# Needs Selected Items: true
# @version 2.1.0

settings = {
  'load_file' => 'concordance',
  'product' => {
    'type' => 'tiff',
    'options' => {
      'naming' => 'full',
      'path' => 'IMAGES'
    }
  },
  'numbering' => {
    'delimiter' => '',
    'groupDocumentPages' => false,
    'groupFamilyItems' => false,
    'advancedView' => false,
    'folder' => {
      'minWidth' => 6,
      'from' => 0,
      'to' => 999,
      'startAt' => 0
    },
    'page' => {
      'minWidth' => 3,
      'from' => 0,
      'to' => 999,
      'startAt' => 1
    }
  }
}

load File.join(__dir__, 'nx_progress.rb') # v1.0.0
load File.join(__dir__, 'summary_reporter.rb') # v1.0.0

require 'fileutils'

# Class for producing items by custodian.
# * +@export_dir+ is the export directory
# * +@reports_dir+ is the reports path
# * +@items+ are the selected items
# * +@settings+ are the export settings
class CustodianProduction < NxProgress
  # Produces items by custodian and creates summary report.
  #
  # @param export_dir [String] path for exports
  # @param items [Collection<Item>] items to export
  # @param settings [Hash]
  def initialize(export_dir, items, settings)
    @export_dir = export_dir
    @reports_dir = File.join(@export_dir, 'REPORTS')
    @items = items
    @settings = settings
    ProgressDialog.forBlock do |progress_dialog|
      super(progress_dialog, 'Custodian Production')
      run if custodian_check
      close_nx
    end
  end

  # Adds loadfile.dat with handling for headers.
  #
  # @param dat_file [String] path of destination file
  # @param new [String] path of loadfile.dat to add
  def add_dat(dat_file, new)
    # Start by copying the whole file, headers included.
    return FileUtils.cp(new, dat_file) unless File.exist?(dat_file)

    f = File.new(dat_file, 'a')
    # Write each line, skipping the header row.
    File.open(new, 'r').each_with_index { |l, i| f.write(l) unless i.zero? }
    f.close
  end

  # Creates BatchExporter with load file and product from @settings.
  #
  # @return [BatchExporter]
  def create_exporter
    exporter = $utilities.create_batch_exporter(@export_dir)
    exporter.add_load_file(@settings['load_file'])
    p = @settings['product']
    exporter.add_product(p['type'], p['options'])
    exporter
  end

  # Creates BatchExporter for populating PDF store.
  #
  # @param dir [String] temp directory
  # @return [BatchExporter]
  def create_populator(dir)
    populator = $utilities.create_batch_exporter(dir)
    populator.add_product('pdf', 'naming' => 'md5')
    populator.set_numbering_options('createProductionSet' => false)
    populator
  end

  # Creates production set for custodian, adds items, and renumbers.
  #
  # @param custodian [String]
  # @return [ProductionSet]
  def create_production(custodian)
    i = intersect_items("custodian:\"#{custodian}\"")
    @@dialog.logMessage("#{custodian} has #{i.size} items")
    return nil if i.empty?

    @@dialog.setSubStatusAndLogIt("Creating Production Set #{custodian}")
    p_set = create_set(custodian)
    p_set.add_items(i)
    p_set.renumber('sortOrder' => 'position')
    @@progress += i.size
    p_set
  end

  # Creates production set for custodian.
  #
  # @param custodian [String]
  # @return [ProductionSet]
  def create_set(custodian)
    p_set = $current_case.new_production_set(custodian)
    opts = @settings['numbering']
    opts['prefix'] = custodian
    p_set.set_numbering_options(opts)
    p_set
  end

  # Checks to ensure all export items have a custodian.
  #
  # @return [true, false] if all items for export have a custodian
  def custodian_check
    i = intersect_items('has-custodian:0')
    return true if i.empty?

    @@dialog.setMainStatusAndLogIt('ERROR - Items missing custodian')
    i.each { |item| puts @@dialog.logMessage(item.get_guid) }
    false
  end

  # Exports items after creating production sets for each custodian.
  #  Moves reports after each export.
  def export
    exporter = create_exporter
    @@dialog.setSubProgress(0, @items.size)
    $current_case.get_all_custodians.each do |c|
      break if @@dialog.abortWasRequested

      production_set = create_production(c)
      next if production_set.nil? || @@dialog.abortWasRequested

      @@dialog.setSubStatusAndLogIt("Exporting Production Set #{c}")
      exporter.export_items(production_set)
      move_reports(c)
    end
  end

  # Logs the current progress and updates the dialog's sub-progress.
  def export_progress
    @@dialog.logMessage("Exported #{@@progress} of #{@items.size} items")
    @@dialog.setSubProgress(@@progress)
  end

  # Returns items from @items that match the query.
  #
  # @param query [String] Nuix query string
  # @return [Collection<Item>]
  def intersect_items(query)
    i = $current_case.search_unsorted(query)
    $utilities.get_item_utility.intersection(i, @items)
  end

  # Moves files from the export directory to custodian report directory.
  #
  # @param custodian [String] the custodian name
  def move_reports(custodian)
    report_dir = File.join(@reports_dir, custodian)
    @@dialog.setSubStatusAndLogIt("Moving reports to #{report_dir}")
    FileUtils.mkdir_p(report_dir)
    Dir.glob(File.join(@export_dir, '*.*')).each do |f|
      @@dialog.logMessage("Moving: #{f}")
      File.rename(f, File.join(report_dir, File.basename(f)))
    end
    export_progress
  end

  # Populates PDF stores for items without stored PDF.
  #
  # @return [true, nil] if PDFs were populated, nil if not needed
  def populate
    @@dialog.logMessage("#{@items.size} selected items")
    i = intersect_items('-has-stored:pdf')
    @@dialog.logMessage("#{i.size} selected items without stored PDF")
    return nil if i.empty?

    @@dialog.setMainStatusAndLogIt('Populating PDF Store')
    populate_pdf(i)
  end

  # Exports PDFs of items to temp directory and then removes temp.
  #
  # @return [true] if PDFs were populated
  def populate_pdf(items)
    @@dialog.setSubStatusAndLogIt("Generating PDFs for #{items.size} items")
    temp_dir = File.join(@export_dir, 'stores')
    @@dialog.logMessage("Using temp directory: #{temp_dir}")
    FileUtils.mkdir_p(temp_dir)
    create_populator(temp_dir).export_items(items)
    @@dialog.setSubStatusAndLogIt("Removing #{temp_dir}")
    FileUtils.rm_r(temp_dir)
    true
  end

  # Populates PDF stores (if required), then exports and summarizes.
    def run
    @@dialog.setMainProgress(0, 3)
    start = Time.now
    @@dialog.setMainProgress(1) if populate
    return nil if @@dialog.abortWasRequested

    @@dialog.setMainStatusAndLogIt("Exporting to #{@export_dir}")
    export
    return nil if @@dialog.abortWasRequested

    @@dialog.setMainProgress(2)
    summarize
    SummaryReporter.new(start, @export_dir, @reports_dir, 'Custodian').write
  end

  # Summarizes the per-custodian reports.
  def summarize
    @@dialog.setMainStatusAndLogIt('Summarizing reports')
    @@dialog.setSubProgress(0, 4)
    # files that are appended
    summarize_append
    @@dialog.setSubProgress(2)
    @@dialog.setSubStatusAndLogIt('Generating loadfile.dat')
    # DAT only needs headers once
    summarize_dat
    @@dialog.setSubProgress(3)
  end

  # Appends loadfile.opt and top-level-MD5-digests.txt from the exports.
  def summarize_append
    ['loadfile.opt', 'top-level-MD5-digests.txt'].each_with_index do |n, i|
      @@dialog.setSubStatusAndLogIt("Generating #{n}")
      @@dialog.setSubProgress(i)
      Dir.glob(File.join(@reports_dir, '*', n)).each do |report|
        @@dialog.logMessage("Adding #{report}")
        f = File.new(File.join(@export_dir, n), 'a')
        f.write(File.read(report))
        f.close
      end
    end
  end

  # Appends loadfile.dat files from the exports, with handling for headers.
  def summarize_dat
    dat_file = File.join(@export_dir, 'loadfile.dat')
    Dir.glob(File.join(@reports_dir, '*', 'loadfile.dat')).each do |r|
      @@dialog.logMessage("Adding #{r}")
      add_dat(dat_file, r)
    end
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
  d = choose_export_dir
  CustodianProduction.new(d, $current_selected_items, settings) unless d.nil?
end
