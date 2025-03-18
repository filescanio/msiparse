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
    """Save digital signatures from the MSI file to disk"""
    if not parent.msi_file_path:
        parent.show_error("Error", "No MSI file selected")
        return
        
    # Get output directory
    output_dir = parent.get_output_directory()
    if not output_dir:
        return
    
    # Show progress
    parent.progress_bar.setVisible(True)
    parent.statusBar().showMessage("Saving digital signatures to disk...")
    
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
    
    if "MSI file has a digital signature" in output:
        parent.statusBar().showMessage("Digital signature found and saved successfully")
        
        if "Successfully extracted" in output:
            extracted_files = [line.split(" to ")[1].strip() for line in output.split('\n') 
                             if "Successfully extracted" in line and " to " in line]
            
            parent.extracted_cert_files = extracted_files
            QMessageBox.information(
                parent,
                "Save Complete",
                "Digital signatures have been saved successfully."
            )
        else:
            parent.statusBar().showMessage("Signature found but extraction may have failed")
            parent.show_warning("Extraction Issue", "Signature found but extraction may have failed. Check the output directory.")
    elif "MSI file does not have a digital signature" in output:
        parent.statusBar().showMessage("No digital signature found in the MSI file")
        QMessageBox.information(
            parent,
            "No Signature",
            "This MSI file does not contain a digital signature."
        )
    else:
        parent.statusBar().showMessage("Unexpected output from certificate extraction")
        parent.show_warning("Extraction Issue", "Unexpected output from certificate extraction. Check the log for details.")
        
def analyze_certificate(parent, show_dialogs=False):
    """Analyze the extracted certificate
    
    Args:
        parent: The parent window object
        show_dialogs: Whether to show message boxes on errors/warnings (default: False)
    """
    if not parent.msi_file_path:
        if show_dialogs:
            parent.show_warning("No MSI File", "Please select an MSI file first.")
        return
        
    # Check if certificate analysis libraries are available
    if not CERT_ANALYSIS_AVAILABLE:
        if show_dialogs:
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
        # Show progress
        parent.progress_bar.setVisible(True)
        parent.statusBar().showMessage("Extracting digital signatures to temporary location...")
        
        # Create a temporary directory for extraction
        with temp_directory() as temp_dir:
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
                    parent.statusBar().showMessage("Digital signature found, analyzing...")
                    
                    # Find extracted files
                    extracted_files = [line.split(" to ")[1].strip() for line in output.split('\n') 
                                     if "Successfully extracted" in line and " to " in line]
                    
                    # If no files were extracted, show warning and return
                    if not extracted_files:
                        parent.statusBar().showMessage("No signature files were extracted")
                        parent.progress_bar.setVisible(False)
                        return
                        
                    # Analyze the extracted certificates
                    _analyze_certificate_files(parent, extracted_files)
                    
                elif "MSI file does not have a digital signature" in output:
                    parent.statusBar().showMessage("No digital signature found in the MSI file")
                    parent.progress_bar.setVisible(False)
                    
                    # Only show message box if explicitly requested
                    if show_dialogs:
                        QMessageBox.information(
                            parent,
                            "No Signature",
                            "This MSI file does not contain a digital signature."
                        )
                else:
                    parent.statusBar().showMessage("Unexpected output from certificate extraction")
                    parent.progress_bar.setVisible(False)
                    
                    # Only show warning if explicitly requested
                    if show_dialogs:
                        parent.show_warning("Extraction Issue", "Unexpected output from certificate extraction.")
                    
            except Exception as e:
                parent.statusBar().showMessage(f"Error extracting certificates: {str(e)}")
                parent.progress_bar.setVisible(False)
                
                # Only show error dialog if explicitly requested
                if show_dialogs:
                    parent.show_error("Extraction Error", f"Failed to extract certificates: {str(e)}")
                
            finally:
                # Hide progress bar
                parent.progress_bar.setVisible(False)
    else:
        # Use previously extracted certificate files
        _analyze_certificate_files(parent, parent.extracted_cert_files)
    
    # Scroll to the beginning after analysis is complete
    parent.cert_details.moveCursor(parent.cert_details.textCursor().Start)
        
def _analyze_certificate_files(parent, certificate_files):
    """Internal method to analyze certificate files"""
    # Find the DigitalSignature file (main signature)
    signature_file = next((f for f in certificate_files if os.path.basename(f) == "DigitalSignature"), None)
            
    if not signature_file:
        parent.statusBar().showMessage("Missing DigitalSignature file")
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
        digest_algorithms = [digest_algo['algorithm'].native for digest_algo in signed_data['digest_algorithms']]
        parent.cert_details.append(f"<b>Digest Algorithm:</b> {', '.join(digest_algorithms)}")
        
        # Use cryptography library for certificate analysis
        try:
            pkcs7_obj = pkcs7.load_der_pkcs7_certificates(signature_data)
            if pkcs7_obj:
                parent.cert_details.append("<h3>Certificate Chain</h3>")
                analyze_certificate_chain_simple(parent, pkcs7_obj)
        except Exception as e:
            parent.statusBar().showMessage(f"Error parsing certificates: {str(e)}")
            parent.cert_details.append(f"<p style='color:red'>Error parsing certificates: {str(e)}</p>")
            
        # Try to extract signer information
        try:
            parent.cert_details.append("<h3>Signer Information</h3>")
            analyze_signer_info_simple(parent, signed_data)
        except Exception as e:
            parent.statusBar().showMessage(f"Error parsing signer info: {str(e)}")
            parent.cert_details.append(f"<p style='color:red'>Error parsing signer info: {str(e)}</p>")
        
        # Signal success in status bar
        parent.statusBar().showMessage("Certificate analysis completed successfully")
        
        # Scroll to the beginning
        parent.cert_details.moveCursor(parent.cert_details.textCursor().Start)
            
    except Exception as e:
        parent.statusBar().showMessage(f"Error analyzing certificate: {str(e)}")
        parent.cert_details.append(f"<p style='color:red'>Error analyzing certificate: {str(e)}</p>")
        parent.cert_details.moveCursor(parent.cert_details.textCursor().Start)
        
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
    name_components = {
        NameOID.COMMON_NAME: 'CN',
        NameOID.ORGANIZATION_NAME: 'O',
        NameOID.ORGANIZATIONAL_UNIT_NAME: 'OU',
        NameOID.COUNTRY_NAME: 'C',
        NameOID.STATE_OR_PROVINCE_NAME: 'ST',
        NameOID.LOCALITY_NAME: 'L'
    }
    
    result = []
    for attr in name:
        prefix = name_components.get(attr.oid, attr.oid._name)
        result.append(f"{prefix}={attr.value}")
    return ", ".join(result) 