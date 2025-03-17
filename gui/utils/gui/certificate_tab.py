"""
Certificate tab functionality for the MSI Parser GUI
"""

import os
import subprocess
import datetime
from PyQt5.QtWidgets import QMessageBox

from utils.common import temp_directory

# Import for certificate analysis
try:
    from cryptography import x509
    from cryptography.hazmat.primitives.serialization import pkcs7
    from cryptography.x509.oid import NameOID, ExtensionOID
    from asn1crypto import cms
    CERT_ANALYSIS_AVAILABLE = True
except ImportError:
    CERT_ANALYSIS_AVAILABLE = False

def extract_certificates(parent):
    """Extract digital signatures from the MSI file"""
    if not parent.msi_file_path:
        parent.show_error("Error", "No MSI file selected")
        return
        
    # Get output directory
    output_dir = parent.get_output_directory()
    if not output_dir:
        return
        
    # Clear previous status
    parent.cert_status.clear()
    parent.cert_status.append("Extracting digital signatures...")
    
    # Show progress
    parent.progress_bar.setVisible(True)
    
    # Build command
    command = [
        parent.msiparse_path,
        "extract_certificate",
        parent.msi_file_path,
        output_dir
    ]
    
    # Run command
    parent.run_command(command, parent.handle_certificate_extraction_complete)
    
def handle_certificate_extraction_complete(parent, output):
    """Handle completion of certificate extraction command"""
    parent.progress_bar.setVisible(False)
    
    # Process the output
    if "MSI file has a digital signature" in output:
        parent.cert_status.append("✅ Digital signature found in the MSI file")
        
        # Check for successful extraction messages
        if "Successfully extracted" in output:
            parent.cert_status.append("\nExtracted signature files:")
            
            # Parse the output to find extracted files
            extracted_files = []
            for line in output.split('\n'):
                if "Successfully extracted" in line:
                    # Extract the file path from the output
                    parts = line.split(" to ")
                    if len(parts) == 2:
                        file_path = parts[1].strip()
                        file_name = os.path.basename(file_path)
                        parent.cert_status.append(f"- {file_name}")
                        extracted_files.append(file_path)
            
            # Store the extracted files for later analysis
            parent.extracted_cert_files = extracted_files
            
            # Enable the analyze button
            parent.analyze_cert_button.setEnabled(True)
            
            # Add a note about what to do with the certificates
            parent.cert_status.append("\nThese files contain the digital signature data. You can:")
            parent.cert_status.append("1. Use tools like 'signtool verify' to validate the signature")
            parent.cert_status.append("2. Extract certificate information with OpenSSL or similar tools")
            parent.cert_status.append("3. Click 'Analyze Signature' to view detailed certificate information")
            
            # Show success message
            QMessageBox.information(
                parent,
                "Extraction Complete",
                "Digital signatures have been extracted successfully."
            )
        else:
            parent.cert_status.append("⚠️ Signature found but extraction may have failed. Check the output directory.")
            parent.analyze_cert_button.setEnabled(False)
    elif "MSI file does not have a digital signature" in output:
        parent.cert_status.append("❌ No digital signature found in the MSI file")
        parent.analyze_cert_button.setEnabled(False)
        
        # Show info message
        QMessageBox.information(
            parent,
            "No Signature",
            "This MSI file does not contain a digital signature."
        )
    else:
        parent.cert_status.append("⚠️ Unexpected output from certificate extraction:")
        parent.cert_status.append(output)
        parent.analyze_cert_button.setEnabled(False)
        
        # Show warning
        parent.show_warning("Extraction Issue", "Unexpected output from certificate extraction. Check the log for details.")
        
def analyze_certificate(parent):
    """Analyze the extracted certificate"""
    if not parent.msi_file_path:
        parent.show_warning("No MSI File", "Please select an MSI file first.")
        return
        
    # Check if certificate analysis libraries are available
    if not CERT_ANALYSIS_AVAILABLE:
        parent.show_warning(
            "Missing Dependencies", 
            "Certificate analysis requires the 'cryptography' and 'asn1crypto' libraries.\n\n"
            "Please install them with:\npip install cryptography asn1crypto"
        )
        return
        
    # Clear previous details
    parent.cert_details.clear()
    
    # If certificates haven't been extracted yet, extract them to a temporary directory
    if not hasattr(parent, 'extracted_cert_files') or not parent.extracted_cert_files:
        parent.cert_status.clear()
        parent.cert_status.append("Extracting digital signatures to temporary location...")
        
        # Create a temporary directory for extraction
        with temp_directory() as temp_dir:
            # Show progress
            parent.progress_bar.setVisible(True)
            
            try:
                # Build and run command
                command = [
                    parent.msiparse_path,
                    "extract_certificate",
                    parent.msi_file_path,
                    temp_dir
                ]
                
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    check=True
                )
                
                output = result.stdout
                
                # Process the output
                if "MSI file has a digital signature" in output:
                    parent.cert_status.append("✅ Digital signature found in the MSI file")
                    
                    # Find extracted files
                    extracted_files = []
                    for line in output.split('\n'):
                        if "Successfully extracted" in line:
                            # Extract the file path from the output
                            parts = line.split(" to ")
                            if len(parts) == 2:
                                file_path = parts[1].strip()
                                file_name = os.path.basename(file_path)
                                parent.cert_status.append(f"- {file_name}")
                                extracted_files.append(file_path)
                    
                    # If no files were extracted, show warning and return
                    if not extracted_files:
                        parent.cert_status.append("⚠️ No signature files were extracted.")
                        parent.progress_bar.setVisible(False)
                        return
                        
                    # Analyze the extracted certificates
                    _analyze_certificate_files(parent, extracted_files)
                    
                elif "MSI file does not have a digital signature" in output:
                    parent.cert_status.append("❌ No digital signature found in the MSI file")
                    parent.progress_bar.setVisible(False)
                    
                    # Show info message
                    QMessageBox.information(
                        parent,
                        "No Signature",
                        "This MSI file does not contain a digital signature."
                    )
                else:
                    parent.cert_status.append("⚠️ Unexpected output from certificate extraction:")
                    parent.cert_status.append(output)
                    parent.progress_bar.setVisible(False)
                    
                    # Show warning
                    parent.show_warning("Extraction Issue", "Unexpected output from certificate extraction. Check the log for details.")
                    
            except Exception as e:
                parent.cert_status.append(f"❌ Error extracting certificates: {str(e)}")
                parent.progress_bar.setVisible(False)
                parent.show_error("Extraction Error", f"Failed to extract certificates: {str(e)}")
                
            finally:
                # Hide progress bar
                parent.progress_bar.setVisible(False)
    else:
        # Use previously extracted certificate files
        _analyze_certificate_files(parent, parent.extracted_cert_files)
        
def _analyze_certificate_files(parent, certificate_files):
    """Internal method to analyze certificate files"""
    # Find the DigitalSignature file (main signature)
    signature_file = None
    for file_path in certificate_files:
        if os.path.basename(file_path) == "DigitalSignature":
            signature_file = file_path
            break
            
    if not signature_file:
        parent.show_warning("Missing Signature File", "The DigitalSignature file was not found among the extracted files.")
        return
        
    try:
        # Read the signature file
        with open(signature_file, 'rb') as f:
            signature_data = f.read()
            
        # Parse the PKCS#7 signature
        content_info = cms.ContentInfo.load(signature_data)
        signed_data = content_info['content']
        
        # Add simple header
        parent.cert_details.append("<h2>Digital Signature Summary</h2>")
        
        # Basic signature information
        parent.cert_details.append("<h3>Signature Information</h3>")
        parent.cert_details.append("<b>Format:</b> DER Encoded PKCS#7 Signed Data")
        
        # Get digest algorithms
        digest_algorithms = []
        for digest_algo in signed_data['digest_algorithms']:
            algo_name = digest_algo['algorithm'].native
            digest_algorithms.append(algo_name)
        parent.cert_details.append(f"<b>Digest Algorithm:</b> {', '.join(digest_algorithms)}")
        
        # Use cryptography library for certificate analysis
        try:
            pkcs7_obj = pkcs7.load_der_pkcs7_certificates(signature_data)
            if pkcs7_obj:
                parent.cert_details.append("<h3>Certificate Chain</h3>")
                analyze_certificate_chain_simple(parent, pkcs7_obj)
        except Exception as e:
            parent.cert_details.append(f"<p style='color:red'>Error parsing certificates: {str(e)}</p>")
            
        # Try to extract signer information
        try:
            parent.cert_details.append("<h3>Signer Information</h3>")
            analyze_signer_info_simple(parent, signed_data)
        except Exception as e:
            parent.cert_details.append(f"<p style='color:red'>Error parsing signer info: {str(e)}</p>")
            
    except Exception as e:
        parent.cert_details.append(f"<p style='color:red'>Error analyzing certificate: {str(e)}</p>")
        
def analyze_certificate_chain_simple(parent, certificates):
    """Analyze the certificate chain from a PKCS#7 signature - simplified version"""
    for i, cert in enumerate(certificates):
        # Determine certificate type
        is_ca = False
        is_self_signed = False
        
        # Check if it's a CA certificate
        try:
            basic_constraints = cert.extensions.get_extension_for_oid(ExtensionOID.BASIC_CONSTRAINTS)
            is_ca = basic_constraints.value.ca
        except x509.ExtensionNotFound:
            pass
            
        # Check if it's self-signed
        try:
            is_self_signed = (cert.subject == cert.issuer)
        except:
            pass
            
        # Determine certificate role
        cert_role = "Unknown"
        if is_self_signed and is_ca:
            cert_role = "Root CA"
        elif is_ca:
            cert_role = "Intermediate CA"
        elif i == 0:
            cert_role = "End Entity (Signer)"
            
        # Get subject and issuer
        subject = get_name_as_text(parent, cert.subject)
        issuer = get_name_as_text(parent, cert.issuer)
        
        # Get validity period using UTC-aware methods to avoid deprecation warnings
        try:
            # Use the new UTC-aware methods if available
            not_before = cert.not_valid_before_utc
            not_after = cert.not_valid_after_utc
        except AttributeError:
            # Fall back to the deprecated methods if the newer ones aren't available
            # (for compatibility with older versions of cryptography)
            not_before = cert.not_valid_before
            not_after = cert.not_valid_after
            
        # Get current time in UTC for comparison
        now = datetime.datetime.now(datetime.timezone.utc)
        is_valid = not_before <= now <= not_after
        
        # Format validity status
        validity_status = "✅ Valid" if is_valid else "❌ Invalid"
        
        # Display certificate information
        parent.cert_details.append(f"<b>Certificate {i+1} ({cert_role}):</b>")
        parent.cert_details.append(f"Subject: {subject}")
        parent.cert_details.append(f"Issuer: {issuer}")
        parent.cert_details.append(f"Validity: {validity_status}")
        parent.cert_details.append(f"Valid from {not_before.strftime('%Y-%m-%d')} to {not_after.strftime('%Y-%m-%d')}")
        
        # Add a separator between certificates
        if i < len(certificates) - 1:
            parent.cert_details.append("<br>")
        
def analyze_signer_info_simple(parent, signed_data):
    """Analyze the signer information from a CMS SignedData object - simplified version"""
    if 'signer_infos' not in signed_data or not signed_data['signer_infos']:
        parent.cert_details.append("<p>No signer information found.</p>")
        return
        
    for i, signer_info in enumerate(signed_data['signer_infos']):
        parent.cert_details.append(f"<b>Signer {i+1}:</b>")
        
        # Get signer identifier
        sid = signer_info['sid']
        if sid.name == 'issuer_and_serial_number':
            issuer = sid.chosen['issuer'].human_friendly
            parent.cert_details.append(f"Issuer: {issuer}")
        else:
            parent.cert_details.append(f"Subject Key Identifier: {sid.chosen.native.hex()[:16]}...")
            
        # Get digest algorithm
        digest_algorithm = signer_info['digest_algorithm']['algorithm'].native
        parent.cert_details.append(f"Digest Algorithm: {digest_algorithm}")
        
        # Check for signing time
        if 'signed_attrs' in signer_info and signer_info['signed_attrs']:
            for attr in signer_info['signed_attrs']:
                attr_type = attr['type'].native
                if attr_type == 'signing_time':
                    signing_time = attr['values'][0].native
                    parent.cert_details.append(f"Signing Time: {signing_time}")
                    break
        
        # Add a separator between signers
        if i < len(signed_data['signer_infos']) - 1:
            parent.cert_details.append("<br>")
            
def get_name_as_text(parent, name):
    """Convert an X.509 name to a readable string"""
    result = []
    for attr in name:
        oid = attr.oid
        if oid == NameOID.COMMON_NAME:
            result.append(f"CN={attr.value}")
        elif oid == NameOID.ORGANIZATION_NAME:
            result.append(f"O={attr.value}")
        elif oid == NameOID.ORGANIZATIONAL_UNIT_NAME:
            result.append(f"OU={attr.value}")
        elif oid == NameOID.COUNTRY_NAME:
            result.append(f"C={attr.value}")
        elif oid == NameOID.STATE_OR_PROVINCE_NAME:
            result.append(f"ST={attr.value}")
        elif oid == NameOID.LOCALITY_NAME:
            result.append(f"L={attr.value}")
        else:
            result.append(f"{oid._name}={attr.value}")
    return ", ".join(result) 