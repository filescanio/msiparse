# 📦 msiparse: The universal MSI inspector

<p align="center">
  <img src="assets/logo.png" width="200" alt="Project Logo"/>
</p>

---

[![CLI Build Status](https://github.com/filescanio/msiparse/actions/workflows/cli_only.yml/badge.svg?branch=master)](https://github.com/filescanio/msiparse/actions/workflows/rust.yml)
[![GUI Build Status](https://github.com/filescanio/msiparse/actions/workflows/full_bundle.yml/badge.svg?branch=master)](https://github.com/filescanio/msiparse/actions/workflows/rust.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Rust](https://img.shields.io/badge/Rust-%23000000.svg?e&logo=rust&logoColor=white)](#)
[![Python](https://img.shields.io/badge/Python-3776AB?logo=python&logoColor=fff)](#)

## 🚀 Overview

`msiparse` is a universal command-line interface (CLI) tool designed to parse and inspect MSI files.

Unlike any other alternatives, this tool features:
- **Cross-Platform Compatibility**: Runs seamlessly on Windows, Linux, and macOS, unlike ANY other alternatives.
- **Comprehensive Extraction**: Capable of Listing and extracting both installed files and embedded tables.
- **Advanced Malware Analysis**: By inspecting the tables and embedded files you can potentially detect malicious CustomAction entries, uncovering even LOLBIN-based threats that might not be part of the installer itself.
- **Automation-Ready**: Outputs JSON for easy integration, and communication - eliminating the need to parse complex ASCII art.

## 🚀 Alternatives

Several alternative tools exist, but they come with notable limitations regarding operating system compatibility or features. Most alternatives are designed specifically for **Windows**, which can restrict their usage across multiple platforms. Below is a list of some popular alternatives, along with their OS dependencies:

- **[7z](https://www.7-zip.org/)** - 🖥️ 🐧 Cross platform, file extraction works great, but no tables/metadata extraction
- **[Orca](https://learn.microsoft.com/en-us/windows/win32/msi/orca-exe)** - 🖥️ Windows only
- **[msitools](https://github.com/GNOME/msitools)** - 🐧 Primarily Linux, may be built on windows, but non-trivial to do so
- **[lessmsi](https://github.com/activescott/lessmsi)** - 🖥️ Windows only
- **[MsiQuery](https://github.com/forderud/MsiQuery)** - 🖥️ Windows only
- **[msidump](https://github.com/mgeeky/msidump)** - 🖥️ Windows only
- **[jsMSIx](https://www.jsware.net/jsware/msicode.html)** - 🖥️ Windows only
- **[MsiAnalyzer](https://github.com/radkum/MsiAnalyzer)** - ❌ Should be cross-platform - interesting, but abandoned.
- **[msi-utils](https://github.com/MSAdministrator/msi-utils)** - ❌ Wrapper around other single-platform tools


## 📂 Usage & Features

```bash
Parse and inspect MSI files

Usage: msiparse [OPTIONS] <COMMAND>

Commands:
  list_metadata        List all the metadata the file has
  list_streams         List all the embedded streams, which can be extracted from the binary
  list_tables          List all the tables and its contents embedded into the msi binary
  extract_all          Extract all the embedded binaries
  extract              Extract a single embedded binary
  extract_certificate  Extract a certificate if it exists in the MSI
  help                 Print this message or the help of the given subcommand(s)

Options:
  -p, --pretty   Pretty-print JSON output
  -h, --help     Print help
  -V, --version  Print version
```

## 🛠 Build

Building is as simple as just issuing:

```bash
git clone https://github.com/filescanio/msiparse
cd msiparse
cargo build --release
```


## 📃 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.txt)

## 📫 Contact

For any questions or feedback, feel free to open an issue or a PR.

## 🙏 Acknowledgements

This project wouldn't have been possible without the incredible work of the [MSI library](https://github.com/mdsteele/rust-msi) by Matthew D. Steele. Huge thanks for providing a solid foundation for this tool!

<br><br>

_Made with  🐍, 🦀 and ❤️_