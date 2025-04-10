use clap::{command, Arg, Command};
use serde::Serialize;
use std::fs::File;
use std::io;
use std::path::Path;
use std::io::{Read, Write};
use time::OffsetDateTime;

pub const DIGITAL_SIGNATURE_STREAM_NAME: &str = "\u{5}DigitalSignature";
pub const MSI_DIGITAL_SIGNATURE_EX_STREAM_NAME: &str = "\u{5}MsiDigitalSignatureEx";

// Helper function to sanitize stream names for file system
fn sanitize_stream_name(name: &str) -> String {
    // Remove control characters and other non-printable characters
    name.chars()
        .filter(|&c| c.is_ascii_graphic() || c.is_ascii_whitespace())
        .collect()
}

// Helper function to extract a stream from a compound file and save it to disk
fn extract_cfb_stream(comp_file: &mut cfb::CompoundFile<File>, stream_name: &str, output_dir: &Path) -> bool {
    match comp_file.open_stream(stream_name) {
        Ok(mut stream) => {
            // Sanitize the stream name for file system
            let sanitized_name = sanitize_stream_name(stream_name);
            let output_path = output_dir.join(&sanitized_name);
            
            match File::create(&output_path) {
                Ok(mut file) => {
                    let mut buffer = Vec::new();
                    if stream.read_to_end(&mut buffer).is_ok() {
                        if file.write_all(&buffer).is_ok() {
                            println!("Successfully extracted {} to {}", 
                                stream_name, 
                                output_path.display());
                            return true;
                        } else {
                            eprintln!("Failed to write {} to file", stream_name);
                        }
                    } else {
                        eprintln!("Failed to read {} stream", stream_name);
                    }
                },
                Err(e) => eprintln!("Failed to create output file: {}", e),
            }
        },
        Err(e) => eprintln!("Failed to open {} stream: {}", stream_name, e),
    }
    false
}

// Dump an MSI stream from a package into a file
// Output is a path, file's name will always be the stream's name
fn dump_stream(stream_name: &str, package: &mut msi::Package<File>, output_dir: &Path) -> bool {
    let stream_opt = package.read_stream(stream_name);
    if stream_opt.is_err() {
        eprintln!("Stream '{}' failed to read, ignoring...", stream_name);
        return false;
    }
    let mut stream = stream_opt.unwrap();
    
    // Sanitize the stream name for file system
    let sanitized_name = sanitize_stream_name(stream_name);
    let stream_path = output_dir.join(&sanitized_name);

    let file_result = File::create(&stream_path);

    if file_result.is_ok() {
        println!("Copying stream '{}' to file '{}'", stream_name, stream_path.to_str().unwrap());
        io::copy(&mut stream, &mut file_result.unwrap()).expect("io::copy failed");
        true
    } else {
        println!("Stream path '{}' was failed to write into, ignoring stream...", stream_path.to_str().unwrap());
        false
    }
}

// CLI main function
// Extract every stream from the package into separate files specified in the output_dir
fn extractall(input: &str, output_dir: &Path) {
    let mut package = msi::open(input).expect("open package");
    let stream_names: Vec<_> = package.streams().collect();

    for stream_name in stream_names {
        dump_stream(stream_name.as_str(), &mut package, output_dir);
    }
}

// CLI main function
// Extract a single stream from the package into the folder specified as the output_dir
fn extract(stream_name: &str, input: &str, output_dir: &Path) {
    let mut package = msi::open(input).expect("open package");
    dump_stream(stream_name, &mut package, output_dir);
}

// CLI main function
// Extract digital signatures from the MSI file using the CFB library
fn extract_certificate(input: &str, output_dir: &Path) {
    match cfb::open(input) {
        Ok(mut comp_file) => {
            let has_signature = comp_file.exists(DIGITAL_SIGNATURE_STREAM_NAME);
            let has_signature_ex = comp_file.exists(MSI_DIGITAL_SIGNATURE_EX_STREAM_NAME);
            
            if has_signature || has_signature_ex {
                println!("MSI file has a digital signature");
                
                // Extract the DigitalSignature stream if it exists
                if has_signature {
                    extract_cfb_stream(&mut comp_file, DIGITAL_SIGNATURE_STREAM_NAME, output_dir);
                }
                
                // Extract the MsiDigitalSignatureEx stream if it exists
                if has_signature_ex {
                    extract_cfb_stream(&mut comp_file, MSI_DIGITAL_SIGNATURE_EX_STREAM_NAME, output_dir);
                }
            } else {
                println!("MSI file does not have a digital signature");
            }
        },
        Err(e) => eprintln!("Failed to open MSI file as a Compound File Binary: {}", e),
    }
}


#[derive(Serialize)]
struct MsiTable {
    name: String,
    columns: Vec<String>,
    rows: Vec<Vec<String>>,
}

// CLI main function
// Dump every table and its contents into a json containing the column headers and all the rows
fn list_tables(input: &str, pretty: bool) {
    let package_iteration = msi::open(input).expect("open package");
    let mut package_queries = msi::open(input).expect("open package");

    let mut tables = Vec::new();
    for table in package_iteration.tables() {
        let rows: Vec<Vec<String>> = package_queries
            .select_rows(msi::Select::table(table.name()))
            .expect("select")
            .map(|row| {
                let mut strings = Vec::with_capacity(row.len());
                for index in 0..row.len() {
                    let string = row[index].to_string().trim_matches('"').to_string();
                    strings.push(string);
                }
                strings
            })
            .collect();

        let columns: Vec<String> = package_queries
            .get_table(table.name())
            .unwrap()
            .columns()
            .iter()
            .map(|column| column.name().to_string())
            .collect();

        tables.push(MsiTable {
            name: table.name().to_string(),
            columns,
            rows,
        });
    }

    let serialized = if pretty {
        serde_json::to_string_pretty(&tables).unwrap()
    } else {
        serde_json::to_string(&tables).unwrap()
    };
    println!("{serialized}");
}

// CLI main function
// List all the extractable streams if someone only wants to extract a single stream (using the 'extract' command)
fn list_streams(input: &str, pretty: bool) {
    let package = msi::open(input).expect("open package");
    let stream_names: Vec<_> = package.streams().collect();

    let serialized = if pretty {
        serde_json::to_string_pretty(&stream_names).unwrap()
    } else {
        serde_json::to_string(&stream_names).unwrap()
    };
    println!("{serialized}")
}

#[derive(Serialize)]
struct MsiMetaData {
    pub title: String,
    pub subject: String,
    pub author: String,
    pub uuid: String,
    pub arch: String,
    pub languages: Vec<String>,
    pub created_at: String,
    pub created_with: String,
    pub is_signed: bool,
    pub codepage: String,
    pub codepage_id: String,
    pub word_count: i32,
    pub comments: String,
}

impl Default for MsiMetaData {
    fn default() -> Self {
        MsiMetaData {
            title: String::default(),
            subject: String::default(),
            author: String::default(),
            uuid: String::default(),
            arch: String::default(),
            languages: Vec::default(),
            created_at: String::default(),
            created_with: String::default(),
            is_signed: false,
            codepage: String::default(),
            codepage_id: String::default(),
            word_count: -1,
            comments: String::default(),
        }
    }
}

// CLI main function
// Get all the metadata that the library is providing us
fn get_metadata(input: &str, pretty: bool) {
    let package = msi::open(input).expect("open package");
    let summary = package.summary_info();

    let mut meta = MsiMetaData::default();

    if let Some(title) = summary.title() {
        meta.title = title.to_string();
    }

    if let Some(subject) = summary.subject() {
        meta.subject = subject.to_string();
    }

    if let Some(author) = summary.author() {
        meta.title = author.to_string();
    }

    if let Some(uuid) = summary.uuid() {
        meta.uuid = uuid.hyphenated().to_string();
    }

    if let Some(arch) = summary.arch() {
        meta.arch = arch.to_string();
    }

    meta.languages = summary
        .languages()
        .iter()
        .map(|lang| lang.tag().to_string())
        .collect();

    if let Some(created_at) = summary.creation_time() {
        meta.created_at = OffsetDateTime::from(created_at).to_string();
    }

    if let Some(created_with) = summary.creating_application() {
        meta.created_with = created_with.to_string();
    }

    if let Some(created_with) = summary.creating_application() {
        meta.created_with = created_with.to_string();
    }

    meta.is_signed = package.has_digital_signature();
    meta.codepage = summary.codepage().name().to_string();
    meta.codepage_id = summary.codepage().id().to_string();

    if let Some(word_count) = summary.word_count() {
        meta.word_count = word_count;
    }

    if let Some(comments) = summary.comments() {
        meta.comments = comments.to_string();
    }

    let serialized = if pretty {
        serde_json::to_string_pretty(&meta).unwrap()
    } else {
        serde_json::to_string(&meta).unwrap()
    };
    println!("{serialized}")
}

// CLI options here we come
// TODO: Add option for pretty/human readable printing
fn main() {
    let matches = command!()
        .propagate_version(true)
        .subcommand_required(true)
        .arg_required_else_help(true)
        .author("OPSWAT, based on the work of Matthew D. Steele <mdsteele@alum.mit.edu>")
        .about("Parse and inspect MSI files")
        .arg(
            Arg::new("pretty")
                .short('p')
                .long("pretty")
                .action(clap::ArgAction::SetTrue)
                .global(true) // Make this flag available to all subcommands
                .help("Pretty-print JSON output"),
        )
        .subcommand(
            Command::new("list_metadata")
                .about("List all the metadata the file has")
                .arg(Arg::new("in_path").required(true)),
        )
        .subcommand(
            Command::new("list_streams")
                .about("List all the embedded streams, which can be extracted from the binary")
                .arg(Arg::new("in_path").required(true)),
        )
        .subcommand(
            Command::new("list_tables")
                .about("List all the tables and its contents embedded into the msi binary")
                .arg(Arg::new("in_path").required(true)),
        )
        .subcommand(
            Command::new("extract_all")
                .about("Extract all the embedded binaries")
                .arg(Arg::new("in_path").required(true))
                .arg(Arg::new("out_folder").required(true)),
        )
        .subcommand(
            Command::new("extract")
                .about("Extract a single embedded binary")
                .arg(Arg::new("in_path").required(true))
                .arg(Arg::new("out_folder").required(true))
                .arg(Arg::new("stream_name").required(true)),
        )
        .subcommand(
            Command::new("extract_certificate")
                .about("Extract a certificate if it exists in the MSI")
                .arg(Arg::new("in_path").required(true))
                .arg(Arg::new("out_folder").required(true)),
        )
        .get_matches();

    let pretty = matches.get_flag("pretty");

    match matches.subcommand() {
        Some(("extract_all", sub_matches)) => extractall(
            sub_matches
                .get_one::<String>("in_path")
                .expect("Path missing"),
            Path::new(
                sub_matches
                    .get_one::<String>("out_folder")
                    .expect("Output missing"),
            ),
        ),
        Some(("extract", sub_matches)) => extract(
            sub_matches
                .get_one::<String>("stream_name")
                .expect("Stream missing"),
            sub_matches
                .get_one::<String>("in_path")
                .expect("Path missing"),
            Path::new(
                sub_matches
                    .get_one::<String>("out_folder")
                    .expect("Output missing"),
            ),
        ),
        Some(("extract_certificate", sub_matches)) => extract_certificate(
            sub_matches
                .get_one::<String>("in_path")
                .expect("Path missing"),
            Path::new(
                    sub_matches
                        .get_one::<String>("out_folder")
                        .expect("Output missing"),
                ),
        ),
        Some(("list_streams", sub_matches)) => list_streams(
            sub_matches
                .get_one::<String>("in_path")
                .expect("Path missing"),
            pretty,
        ),
        Some(("list_tables", sub_matches)) => list_tables(
            sub_matches
                .get_one::<String>("in_path")
                .expect("Path missing"),
            pretty,
        ),
        Some(("list_metadata", sub_matches)) => get_metadata(
            sub_matches
                .get_one::<String>("in_path")
                .expect("Path missing"),
            pretty,
        ),
        _ => unreachable!("Exhausted list of subcommands and subcommand_required prevents `None`"),
    }
}
