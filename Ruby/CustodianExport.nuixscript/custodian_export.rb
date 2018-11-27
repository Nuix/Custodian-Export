# Script for exporting items by custodian.
# @author mrk
# @version 2.0

# Class for Nx Dialog.
# * +@@dialog+ is an Nx ProcessDialog
# * +@@progress+ represents the main progress
class NxClass
  require File.join(__dir__, 'Nx.jar')
  java_import 'com.nuix.nx.NuixConnection'
  java_import 'com.nuix.nx.LookAndFeelHelper'
  java_import 'com.nuix.nx.dialogs.ProgressDialog'
  LookAndFeelHelper.setWindowsIfMetal
  NuixConnection.setUtilities($utilities)
  NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

  # Initializes Progress Dialog.
  #
  # @param progress_dialog [ProgressDialog] the progress dialog
  # @param title [String] the title
  def initialize(progress_dialog, title)
    progress_dialog.setTitle(title)
    progress_dialog.setLogVisible(true)
    progress_dialog.setTimestampLoggedMessages(true)
    @@dialog = progress_dialog
    @@progress = 0
  end

  # Increments and sets main progress.
  def advance_main
    @@progress += 1
    @@dialog.setMainProgress(@@progress)
  end

  # Completes the dialog, or logs the abortion.
  def close_nx
    if @@dialog.abortWasRequested
      @@dialog.logMessage('Aborted')
    else
      @@dialog.setCompleted
    end
  end
end

# Class for exporting items by custodian.
# * +@export_dir+ is the export directory
# * +@path[:exports]+ is the exports path
# * +@path[:reports]+ is the reports path
# * +@path[:summary]+ is the summary report file
class CustodianExport < NxClass
  require 'fileutils'

  # Exports items by custodian and creates summary report.
  def initialize(export_dir, items)
    NxClass::ProgressDialog.forBlock do |progress_dialog|
      super(progress_dialog, 'Custodian Export')
      @export_dir = export_dir
      @path = { exports: File.join(export_dir, 'Export'),
                reports: File.join(export_dir, 'Reports'),
                summary: File.join(export_dir, 'summary-report.xml') }
      @path.each { |k, v| @@dialog.logMessage("Writing #{k} to #{v}") }
      run(export_dir, items)
      close_nx
    end
  end

  protected

  # Exports items.
  #
  # @param items [Set<Item>] Nuix items to export
  def export(items)
    exporter = Exporter.new(items)
    return false if exporter.custodian_missing

    exporter.export(@path[:exports])
  end

  # Moves files to reports directory.
  # Creates custodian directory and logs messages.
  #
  # @param dir [String] directory containg files to move
  def move(dir)
    @@dialog.logMessage("Moving files from #{dir}")
    # Make per-custodian directory
    FileUtils.mkdir_p(File.join(@path[:reports], File.basename(dir)))
    @@dialog.logMessage("Moved: #{move_files(dir).join(', ')}")
  end

  # Moves files to reports directory.
  #
  # @param dir [String] directory containg files to move
  # @return [Array] of file names that were moved
  def move_files(dir)
    moved = []
    Dir.entries(dir).each do |e|
      f = File.join(dir, e)
      if File.file?(f)
        File.rename(f, f.sub(@path[:exports], @path[:reports]))
        moved << e
      end
    end
    moved
  end

  # Moves reports.
  def move_reports
    @@dialog.setMainStatusAndLogIt('Preparing summary report')
    @@dialog.setSubStatusAndLogIt("Moving reports to #{@path[:reports]}.")
    to_move = Dir.glob(File.join(@path[:exports], '*', File::SEPARATOR))
    @@dialog.setSubProgress(0, to_move.size)
    to_move.each_with_index do |dir, i|
      move(dir)
      @@dialog.setSubProgress(i)

      break if @@dialog.abortWasRequested
    end
    advance_main
  end

  # Exports items per-custodian, then moves and summarizes reports.
  #
  # @param export_dir [String] directory for export
  # @param items [Set<Item>] Nuix items to export
  def run(export_dir, items)
    start = Time.now
    export(items)
    return false if @@dialog.abortWasRequested

    move_reports
    return false if @@dialog.abortWasRequested

    write(Reporter.new(start, export_dir, @path[:reports]).xml)
  end

  # Writes pretty formatted XML doc to +@path[:summary]+.
  #
  # @param doc [REXML::Document] XML to write
  def write(doc)
    @@dialog.setSubStatusAndLogIt("Writing #{@path[:exports]}")
    formatter = REXML::Formatters::Pretty.new
    formatter.compact = true
    File.open(@path[:summary], 'w') { |f| f.puts formatter.write(doc.root, '') }
  end

  # Class for exporting items.
  # * +@top_items+ is the top-level items to export
  # * +@total+ is the number of top-level items to export
  class Exporter < NxClass
    # Exports top-level items to path.
    #
    # @param items [Set<Item>] Nuix items to export
    def initialize(items)
      @@dialog.setMainStatusAndLogIt('Getting items to export')
      msg = "Finding top-level items from #{items.size} selected items"
      @@dialog.setSubStatusAndLogIt(msg)
      @@dialog.setSubProgress(0, 2)
      @top_items = $utilities.get_item_utility.find_top_level_items(items, true)
      @total = @top_items.size
      @@dialog.logMessage("#{@total} top-level items to export")
      @@dialog.setMainProgress(0, @total + 3)
      @@dialog.setSubProgress(1)
    end

    # Checks if +@top_items+ all have a custodian.
    # If there are items without custodian, logs GUIDs.
    #
    # @return [true, false] true if all +@export_top+ items have a custodian
    def custodian_missing
      missing_custodian = intersect_top('has-custodian:0')
      return false if missing_custodian.empty?

      @@dialog.setMainStatusAndLogIt('ERROR - Items missing custodian')
      missing_custodian.each { |i| puts @@dialog.logMessage(i.get_guid) }
      true
    end

    # Exports each custodian until all items are exported, or aborted.
    #
    # @param path [String] path for exports
    def export(path)
      @@dialog.setMainStatusAndLogIt("Exporting to #{path}")
      @@dialog.setAbortButtonVisible(true)
      custodians = $current_case.get_all_custodians
      @@dialog.setSubProgress(0, custodians.size)
      custodians.each_with_index do |c, index|
        items = intersect_top("custodian:\"#{c}\"")
        next if items.empty?

        @@dialog.setSubProgress(index)
        break if @@dialog.abortWasRequested || batch_export(path, c, items)
      end
    end

    protected

    # Exports items for custodian.
    #
    # @param path [String] path for exports
    # @param name [String] custodian of the items
    # @param items [Collection<Item>] items to export
    # @return [true, false] if all items have been exported
    def batch_export(path, name, items)
      i = items.size
      @@dialog.setSubStatusAndLogIt("Exporting custodian: #{name} (#{i} items)")
      export_options = { naming: 'item_name_with_path', mailFormat: 'pst' }
      exporter = $utilities.create_batch_exporter(File.join(path, name))
      exporter.add_product('native', export_options)
      exporter.set_numbering_options(createProductionSet: false)
      exporter.export_items(items)
      progress_check(i)
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
  end

  # Class for summary-report.xml
  class Reporter < NxClass
    require 'rexml/document'
    # Initializes summary report and summarizes.
    #
    # @param start_time [Time]  start time of export
    # @param export_dir [String]  export path
    # @param reports_path [String]  path containing reports
    def initialize(start_time, export_dir, reports_path)
      @start_time = start_time
      @export_dir = export_dir
      @stats = { export: Hash.new(0),
                 file: Hash.new(0),
                 mime: Hash.new(0) }
      @total_duration = 0
      @configuration = nil
      @custodians = []
      @@dialog.setSubProgress(0, 3)
      summarize(reports_path)
    end

    # Generates summary report XML document.
    #
    # @return [REXML::Document] XML document
    def xml
      @@dialog.setSubStatusAndLogIt('Generating XML')
      @@dialog.setSubProgress(0, 2)
      doc = REXML::Document.new
      doc.add_element('Nuix', nuix_attributes) << xml_export
      @@dialog.setSubProgress(1)
      doc
    end

    protected

    # Attributes for Export element of summary-report.xml.
    #
    # @return [Hash] export attributes.
    def export_attributes
      end_time = Time.now
      { 'startTime' => @start_time,
        'endTime' => end_time,
        'exportDuration' => @total_duration,
        'processingDuration' => end_time - @start_time }
    end

    # Array of Export XML elements.
    #
    # @return [Array<REXML::Element>] elements for Export
    def export_elements
      [xml_configuration,
       xml_stats('ExportStatistics', @stats[:export]),
       xml_custodians,
       xml_stats('FileStatistics', @stats[:file]),
       xml_throughput,
       xml_mimes]
    end

    # Attributes for Nuix element of summary-report.xml.
    #
    # @return [Hash] Nuix attributes
    def nuix_attributes
      { 'version' => NUIX_VERSION,
        'architecture' => ENV_JAVA['os.arch'] }
    end

    # Summarizes reports and generates summary-report.xml.
    #
    # @param reports_path [String] path containing summary reports
    def summarize(reports_path)
      @@dialog.setSubStatusAndLogIt("Summarizing reports in #{reports_path}")
      reports = Dir.glob(File.join(reports_path, '*', 'summary-report.xml'))
      @@dialog.setSubProgress(0, reports.size)
      reports.each_with_index do |f, i|
        summarize_file(f)
        @@dialog.setSubProgress(i)
      end
      advance_main
    end

    # Summarizes data from a summary-report.xml.
    #
    # @param file_path [String] path to a summary-report.xml file
    def summarize_file(file_path)
      @@dialog.logMessage("Reading #{file_path}")
      report = ReportFile.new(file_path)
      @total_duration += report.duration
      @configuration = report.configuration if @configuration.nil?
      @custodians << report.details
      report.statistics.each do |type, values|
        @stats[type].merge!(values) { |_k, v1, v2| v1 + v2 }
      end
    end

    # ExportConfiguration XML.
    # ExportDirectory text updated with @export_dir.
    #
    # @return [REXML::Element] ExportConfiguration XML
    def xml_configuration
      @configuration.elements['ExportDirectory'].text = @export_dir
      @configuration
    end

    # CustodianDetails XML.
    #
    # @return [REXML::Element] CustodianDetails XML
    def xml_custodians
      c_xml = REXML::Element.new 'CustodianDetails'
      @custodians.each { |attrs| c_xml.add_element('Custodian', attrs) }
      c_xml
    end

    # Export XML.
    #
    # @return [REXML::Element] Export XML
    def xml_export
      e = REXML::Element.new 'Export'
      e.add_attributes(export_attributes)
      export_elements.each { |xml| e << xml }
      e
    end

    # MimeTypeStatistics XML.
    #
    # @return [REXML::Element] MimeTypeStatistics XML
    def xml_mimes
      mime_stats_xml = REXML::Element.new 'MimeTypeStatistics'
      e = mime_stats_xml.add_element 'MimeTypes'
      @stats[:mime].each do |k, v|
        e.add_element('MimeType', 'name' => k, 'count' => v)
      end
      mime_stats_xml
    end

    # Generate statistics XML from hash.
    #
    # @param name [String] name of XML element
    # @param stats [Hash] of counts
    # @return [REXML::Element] for statistics
    def xml_stats(name, stats)
      stats_xml = REXML::Element.new(name)
      stats.each { |k, v| stats_xml.add_element(k).text = v }
      stats_xml
    end

    # ThroughputStatistics XML.
    #
    # @return [REXML::Element] ThroughputStatistics XML
    def xml_throughput
      t = @stats[:file]['NativeFilesExported'] / @total_duration.to_f
      throughput_stats_xml = REXML::Element.new 'ThroughputStatistics'
      throughput_stats_xml.add_element('NativeDocRate').text = t
      throughput_stats_xml
    end

    # Class for parsing summary-report.xml files.
    # @example Get ExportConfiguration
    #  ReportFile.configuration #=> Nuix/Export/ExportConfiguration
    # @example Get Custodian Details
    #  ReportFile.details #=> Hash for CustodianDetails
    # @example Get exportDuration
    #  ReportFile.duration #=> Nuix/Export[exportDuration]
    # @example Get Statistics Hash
    #  ReportFile.statistics[:export] #=> Nuix/Export/ExportStatistics
    #  ReportFile.statistics[:file] #=> Nuix/Export/FileStatistics
    #  ReportFile.statistics[:mime] #=> Nuix/Export/MimeTypeStatistics/MimeTypes
    class ReportFile
      # @return [Integer] exportDuration
      attr_reader :duration
      # @return [Hash{
      #  :export => ExportStatistics,
      #  :file => FileStatistics,
      #  :mime => MimeTypeStatistics/MimeTypes }]
      attr_reader :statistics

      # Loads summary-report.xml file to parse.
      #
      # @param file_path [String] path to a summary-report.xml
      def initialize(file_path)
        @custodian = File.basename(File.dirname(file_path))
        doc = REXML::Document.new(IO.read(file_path))
        @export_doc = doc.elements['Nuix/Export']
        @duration = @export_doc.attributes['exportDuration'].to_i
        @statistics = {}
        @statistics[:export] = export_statistics
        @statistics[:file] = file_statistics
        @statistics[:mime] = mimes
      end

      # ExportConfiguration XML.
      #
      # @return [REXML::Element] ExportConfiguration XML
      def configuration
        @export_doc.elements['ExportConfiguration']
      end

      # Creates Hash of details for CustodianDetails.
      #  Includes custodian, duration, and the ExportStatistics.
      #
      # @return [Hash] CustodianDetails
      def details
        d = {}
        d['name'] = @custodian
        d['exportDuration'] = @duration
        @statistics[:export].each do |k, v|
          # fix case
          n = String.new(k)
          n[0] = n[0].downcase
          d[n] = v
        end
        d
      end

      protected

      # ExportStatistics information from XML.
      #
      # @return [Hash] ExportStatistics
      def export_statistics
        fields = %w[SelectedItems ExcludedCount TotalItemsToExport FailedItems]
        stats = {}
        fields.each do |v|
          stats[v] = @export_doc.elements["ExportStatistics/#{v}"].text.to_i
        end
        stats
      end

      # FileStatistics information from XML.
      #
      # @return [Hash] FileStatistics
      def file_statistics
        file_stats = {}
        @export_doc.elements['FileStatistics'].each do |e|
          next unless e.is_a?(REXML::Element)

          file_stats[e.name] = e.text.to_i
        end
        file_stats
      end

      # Creates Hash of MimeTypes counts from XML.
      #
      # @return [Hash] MimeTypes
      def mimes
        types = {}
        @export_doc.elements['MimeTypeStatistics/MimeTypes'].each do |e|
          next unless e.is_a?(REXML::Element)

          types[e.attributes['name']] = e.attributes['count'].to_i
        end
        types
      end
    end
  end
end
