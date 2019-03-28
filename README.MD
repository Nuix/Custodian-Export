Custodian Export
==============

![Last tested in Nuix 7.8](https://img.shields.io/badge/Nuix-7.8-green.svg)

View the GitHub project [here](https://github.com/Nuix/Custodian-Export) or download the latest release [here](https://github.com/Nuix/Custodian-Export/releases).

# Overview

This script exports the selected items by performing a series of custodian based exports.

Custodian Production creates separate production sets for each custodian and exports, using settings from a JSON file (i.e. separate Bates prefixes for each custodian).

Custodian PST Export generates separate PST exports for each custodian's items.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

## Usage

Before running the script you will need to have a case open in Nuix and a selection of items which you wish to work with.

The script will prompt the user to select a directory to be used as the export destination. The items will be exported by custodian, with separate directories for exports and reports.

A new summary-report.xml is created in the export destination, summarizing the separate per-custodian exports. The summary report.xml file contain details of the export operation, including:
- The exact legal export configuration.
- Detailed breakdowns of all files that were exported.
- Timing information for each of the custodian exports.
- Detailed file type statistics.
- Details of all duplicate top-level items not exported.
- A fully qualified query string that can be used to find all items that failed to export correctly.

## Custodian Production

This script creates and exports a separate production set for each custodian, numbered with the custodian name as the prefix. If any of the items do not have an assigned custodian, the export is canceled.

If selected items do not have a stored PDF, the script exports PDFs of the items to a temp directory to populate the store. This enables page-based numbering for the production sets.

### Settings JSON

The settings file (production.json by default) contains settings for the load file, export product, and numbering. The numbering prefix will become the custodian name, so it does not need to be specified.

### Populate PDF Store

If any of the selected items do not have a stored PDF, the script generates images for those items by exporting PDFs to a temp directory. The temp directory is deleted once the export is complete.

Ensuring all the items are imaged enables page-based numbering for the production sets.

### Production Sets

The production sets use numbering loaded from the JSON settings file, with the prefix set to the custodian name. After the selected items for the custodian are added, the production set is renumbered using sortOrder = position.

### Export Product

The production sets are exported using the load file and product options from the JSON settings file.

### Reports

After each per-custodian export completes, the report files (i.e. summary-report.txt/xml, top-level-MD5-digests.txt, and loadfile.opt/dat) are moved to the reports directory into separate directories for each custodian.

After all the exports are complete, the script creates its own report files, combining the individual exports.

## Custodian PST Export

This script exports the top level item for each of the selected items. Items that have no top-level item will also be included. Multiple items which resolve to the same top-level item will result in a single item in the export, so the exported items may be smaller than the number of selected items.

Each custodian is exported separately into folders named for the custodian. If any of the items do not have an assigned custodian, the export is canceled.

### Export Product

Each export includes natives in PST mail format. By default the script will generate PSTs with a flat structure, but if line 7 of custodian_pst.rb is un-commented (i.e. naming = "item_name_with_path") it will be set to re-create directory structure.

### Reports

After the per-custodian exports complete, all their output (i.e. summary-report.txt/xml and top-level-MD5-digests.txt) is moved to the reports directory (maintaining the custodian folders).

## Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the scripts.

# License

```
Copyright 2019 Nuix

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
