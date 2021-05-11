# Class for handling summary-report.xml files
# @version 1.0.1
class SummaryReporter < NxProgress
  require 'rexml/document'

  # Initializes summary report and summarizes.
  #
  # @param start_time [Time]  start time of export
  # @param export_dir [String]  export path
  # @param reports_path [String]  path containing reports
  # @param exports_name [String] details type for each export
  def initialize(start_time, export_dir, reports_path, exports_type)
    @start_time = start_time
    @export_dir = export_dir
    @stats = { export: Hash.new(0), file: Hash.new(0), mime: Hash.new(0) }
    @total_duration = 0
    @configuration = nil
    @exports = []
    @exports_type = exports_type
    @@dialog.setSubProgress(0, 3)
    summarize(reports_path)
  end

  # Writes pretty formatted XML doc.
  #
  # @param file_path [string] path for created summary-report.xml
  def write
    file_path = File.join(@export_dir, 'summary-report.xml')
    @@dialog.setSubStatusAndLogIt("Writing #{file_path}")
    formatter = REXML::Formatters::Pretty.new
    formatter.compact = true
    File.open(file_path, 'w') { |f| f.puts formatter.write(xml.root, '') }
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
     xml_exports,
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
    begin
      r = ReportFile.new(file_path)
    rescue
      @@dialog.logMessage("ERROR loading #{file_path}. Is it XML?")
    else
      @total_duration += r.duration
      @configuration = r.configuration if @configuration.nil?
      @exports << r.details
      r.statistics.each { |t, v| @stats[t].merge!(v) { |_k, v1, v2| v1 + v2 } }
    end
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

  # ExportConfiguration XML.
  # ExportDirectory text updated with @export_dir.
  #
  # @return [REXML::Element] ExportConfiguration XML
  def xml_configuration
    @configuration.elements['ExportDirectory'].text = @export_dir
    @configuration
  end

  # ExportsDetails XML.
  #
  # @return [REXML::Element] ExportsDetails XML
  def xml_exports
    e_xml = REXML::Element.new "#{@exports_type}Details"
    @exports.each { |attrs| e_xml.add_element(@exports_type, attrs) }
    e_xml
  end

  # Export XML.
  #
  # @return [REXML::Element] Export XML
  def xml_export
    e = REXML::Element.new 'Export'
    e.add_attributes(export_attributes)
    begin
      export_elements.each { |xml| e << xml }
    rescue => error
      e.add_element("ERROR: #{error.message}")
      @@dialog.logMessage('ERROR generating summary-report.xml')
    end
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
