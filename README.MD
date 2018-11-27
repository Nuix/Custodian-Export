Custodian Export
==============

![Last tested in Nuix 7.6](https://img.shields.io/badge/Nuix-7.6-green.svg)

View the GitHub project [here](https://github.com/Nuix/Custodian-Export) or download the latest release [here](https://github.com/Nuix/Custodian-Export/releases).

View the Docs [here](https://nuix.github.io/Custodian-Export/).

# Overview

**Written By:** mrk

This script exports the selected items by performing a series of custodian based exports. Each export includes natives, using item name with path, and exporting all messages to a single PST file. After separately exporting each custodian's items, the script moves the reports to their own directory and generates a summary-report.xml of the separate exports.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

## Usage

Before running the script you will need to have a case open in Nuix and a selection of items which you wish to work with.

The script will prompt the user to select a directory to be used as the export destination. Within this export destination, the script will create a summary-report.xml file, and 'Export' and 'Reports' directories, each containing separate folders for each custodian's export and report.

### Export Items

This script exports the top level item for each of the selected items. Items that have no top-level item will also be included. Multiple items which resolve to the same top-level item will result in a single item in the export, so the exported items may be smaller than the number of selected items.

Each custodian is exported separately into folders named for the custodian. If any of the items do not have an assigned custodian, the export is canceled.

### Export Product

Each export includes natives, named using the item name with path, and PST mail format.

### Reports

After the per-custodian exports complete, all their output (i.e. summary-report.txt/xml and top-level-MD5-digests.txt) is moved to the reports directory (maintaining the custodian folders).

A new summary-report.xml is created in the export destination, summarizing the separate per-custodian exports. The summary report.xml file contain details of the export operation, including:
- The exact legal export configuration.
- Detailed breakdowns of all files that were exported.
- Timing information for each of the custodian exports.
- Detailed file type statistics.
- Details of all duplicate top-level items not exported.
- A fully qualified query string that can be used to find all items that failed to export correctly.

After exporting items to a legal load file format, review the associated summary report to ensure the content of the export meets expectations.

# License

```
Copyright 2018 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
