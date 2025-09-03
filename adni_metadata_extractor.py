#!/usr/bin/env python3
"""
ADNI Metadata Detailed Extraction Script

This script extracts detailed metadata from ADNI XML files, organizes them
by scan type into separate Excel sheets, and creates a comprehensive report
with patient metadata arranged by rows.
"""

import argparse
import logging
import os
import re
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class ADNIDetailedMetadataExtractor:
    """ADNI metadata detailed extractor for XML files."""
    
    def __init__(self, base_path="/Users/username/Documents/adni_metainfo"):
        """
        Initialize the ADNI metadata extractor.
        
        Args:
            base_path (str): Base directory path containing ADNI metadata
        """
        self.base_path = Path(base_path)
        self.scan_type_data = defaultdict(list)
        self.namespace = None  # No namespace usage
        
    def extract_scan_type(self, xml_file_path):
        """
        Extract scan type from XML file path/name.
        
        Args:
            xml_file_path (str): Path to the XML file
            
        Returns:
            str: Scan type identifier
        """
        filename = os.path.basename(xml_file_path)
        
        # PET scan type matching
        if "FDG" in filename:
            return "PET_FDG"
        elif "FBB" in filename or "Florbetaben" in filename:
            return "PET_FBB"  
        elif "AV45" in filename or "florbetapir" in filename:
            return "PET_AV45"
        elif ("Tau" in filename or "AV1451" in filename or 
              "FLORTAUCIPIR" in filename):
            return "PET_TAU"
        elif "PET" in filename:
            return "PET_OTHER"
        
        # MRI scan type matching
        elif "MPR" in filename:
            if "FLAIR" in filename:
                return "MRI_FLAIR"
            else:
                return "MRI_MPRAGE"
        elif "FLAIR" in filename:
            return "MRI_FLAIR"
        elif "DTI" in filename:
            return "MRI_DTI"
        elif "rsfMRI" in filename or "fMRI" in filename:
            return "MRI_fMRI"
        elif "ASL" in filename:
            return "MRI_ASL"
        elif "T2" in filename:
            return "MRI_T2"
        else:
            return "OTHER"
    
    def parse_xml_metadata(self, xml_file_path):
        """
        Parse metadata from XML file.
        
        Args:
            xml_file_path (Path): Path to XML file
            
        Returns:
            dict: Extracted metadata or None if parsing fails
        """
        try:
            tree = ET.parse(xml_file_path)
            root = tree.getroot()
            
            metadata = {}
            
            # Extract basic information
            metadata['filename'] = os.path.basename(xml_file_path)
            metadata['scan_type'] = self.extract_scan_type(xml_file_path)
            
            # Subject information
            subject_id = root.find('.//subjectIdentifier')
            metadata['subject_id'] = (subject_id.text 
                                    if subject_id is not None else 'N/A')
            
            research_group = root.find('.//researchGroup')
            metadata['research_group'] = (research_group.text 
                                        if research_group is not None 
                                        else 'N/A')
            
            subject_sex = root.find('.//subjectSex')
            metadata['gender'] = (subject_sex.text 
                                if subject_sex is not None else 'N/A')
            
            subject_age = root.find('.//subjectAge')
            metadata['age'] = (float(subject_age.text) 
                             if subject_age is not None else None)
            
            weight = root.find('.//weightKg')
            metadata['weight_kg'] = (float(weight.text) 
                                   if weight is not None else None)
            
            # APOE genotype information
            apoe_elements = root.findall('.//subjectInfo[@item]')
            for apoe in apoe_elements:
                item = apoe.get('item')
                if 'APOE A1' in item:
                    metadata['apoe_a1'] = apoe.text
                elif 'APOE A2' in item:
                    metadata['apoe_a2'] = apoe.text
            
            # Visit information
            visit_id = root.find('.//visitIdentifier')
            metadata['visit_type'] = (visit_id.text 
                                    if visit_id is not None else 'N/A')
            
            # Scan information
            modality = root.find('.//modality')
            metadata['modality'] = (modality.text 
                                  if modality is not None else 'N/A')
            
            date_acquired = root.find('.//dateAcquired')
            metadata['scan_date'] = (date_acquired.text 
                                   if date_acquired is not None else 'N/A')
            
            series_id = root.find('.//seriesIdentifier')
            metadata['series_id'] = (series_id.text 
                                   if series_id is not None else 'N/A')
            
            # Extract clinical assessment scores
            self._extract_clinical_scores(root, metadata)
            
            # Extract imaging protocol information
            self._extract_imaging_protocol(root, metadata)
            
            # Extract processing information (for MRI)
            self._extract_processing_info(root, metadata)
            
            return metadata
            
        except Exception as e:
            logger.error(f"XML parsing error ({xml_file_path}): {e}")
            return None
    
    def _extract_clinical_scores(self, root, metadata):
        """
        Extract clinical assessment scores from XML.
        
        Args:
            root: XML root element
            metadata (dict): Dictionary to store extracted metadata
        """
        assessments = root.findall('.//assessment')
        
        for assessment in assessments:
            name = assessment.get('name', '')
            
            if 'MMSE' in name:
                score = assessment.find('.//assessmentScore[@attribute="MMSCORE"]')
                metadata['mmse_score'] = (float(score.text) 
                                        if score is not None else None)
                
            elif 'CDR' in name:
                score = assessment.find('.//assessmentScore[@attribute="CDGLOBAL"]')
                metadata['cdr_score'] = (float(score.text) 
                                       if score is not None else None)
                
            elif 'NPI' in name:
                score = assessment.find('.//assessmentScore[@attribute="NPISCORE"]')
                metadata['npi_score'] = (float(score.text) 
                                       if score is not None else None)
                
            elif 'FAQ' in name:
                score = assessment.find('.//assessmentScore[@attribute="FAQTOTAL"]')
                metadata['faq_score'] = (float(score.text) 
                                       if score is not None else None)
    
    def _extract_imaging_protocol(self, root, metadata):
        """
        Extract imaging protocol information from XML.
        
        Args:
            root: XML root element
            metadata (dict): Dictionary to store extracted metadata
        """
        # MRI protocol information
        protocol_terms = root.findall('.//protocolTerm/protocol')
        
        for protocol in protocol_terms:
            term = protocol.get('term', '')
            
            if term == 'TE':
                metadata['te_ms'] = (float(protocol.text) 
                                   if protocol.text else None)
            elif term == 'TR':
                metadata['tr_ms'] = (float(protocol.text) 
                                   if protocol.text else None)
            elif term == 'Slice Thickness':
                metadata['slice_thickness_mm'] = (float(protocol.text) 
                                                if protocol.text else None)
            elif term == 'Flip Angle':
                metadata['flip_angle'] = (float(protocol.text) 
                                        if protocol.text else None)
            elif term == 'Manufacturer':
                metadata['manufacturer'] = protocol.text
            elif term == 'Mfg Model':
                metadata['device_model'] = protocol.text
            elif term == 'Field Strength':
                metadata['field_strength_t'] = (float(protocol.text) 
                                               if protocol.text else None)
        
        # PET specific protocol information
        pet_protocols = root.findall('.//imagingProtocol//protocol')
        for protocol in pet_protocols:
            term = protocol.get('term', '')
            
            if term == 'Radiopharmaceutical':
                metadata['radiopharmaceutical'] = protocol.text
            elif term == 'Number of Rows':
                metadata['num_rows'] = (int(float(protocol.text)) 
                                      if protocol.text else None)
            elif term == 'Number of Columns':
                metadata['num_columns'] = (int(float(protocol.text)) 
                                         if protocol.text else None)
            elif term == 'Number of Slices':
                metadata['num_slices'] = (int(float(protocol.text)) 
                                        if protocol.text else None)
            elif term == 'Pixel Spacing X':
                metadata['pixel_spacing_x'] = (float(protocol.text) 
                                             if protocol.text else None)
            elif term == 'Pixel Spacing Y':
                metadata['pixel_spacing_y'] = (float(protocol.text) 
                                             if protocol.text else None)
            elif term == 'Reconstruction':
                metadata['reconstruction_method'] = protocol.text
    
    def _extract_processing_info(self, root, metadata):
        """
        Extract processing information from XML.
        
        Args:
            root: XML root element
            metadata (dict): Dictionary to store extracted metadata
        """
        processed_label = root.find('.//processedDataLabel')
        if processed_label is not None:
            metadata['processing_label'] = processed_label.text
        
        # Processing steps
        provenance_details = root.findall('.//provenanceDetail')
        processing_steps = []
        
        for detail in provenance_details:
            process = detail.find('.//process')
            program = detail.find('.//program')
            
            if process is not None and program is not None:
                processing_steps.append(f"{process.text}({program.text})")
        
        metadata['processing_steps'] = ('; '.join(processing_steps) 
                                      if processing_steps else 'N/A')
    
    def _find_xml_folders(self, base_dir):
        """
        Recursively find folders containing XML files.
        
        Args:
            base_dir (Path): Base directory to search
            
        Returns:
            list: List of Path objects containing XML files
        """
        xml_folders = []
        base_path = Path(base_dir)
        
        if not base_path.exists():
            return xml_folders
        
        try:
            for root, dirs, files in os.walk(base_path):
                xml_files = [f for f in files if f.endswith('.xml')]
                if xml_files:
                    xml_folders.append(Path(root))
            
            return xml_folders
        except Exception as e:
            logger.warning(f"Error during folder exploration: {e}")
            return xml_folders
    
    def _resolve_folder_path(self, folder_input):
        """
        Convert user input to actual folder path.
        
        Args:
            folder_input (str): User input folder path
            
        Returns:
            Path: Resolved folder path
        """
        folder_path = Path(folder_input)
        
        # Use absolute path as-is
        if folder_path.is_absolute():
            return folder_path
        
        # Interpret relative path based on base_path
        resolved_path = self.base_path / folder_path
        
        # Check if path exists
        if resolved_path.exists():
            return resolved_path
        
        # Check if /ADNI subfolder exists (ADNI metadata structure)
        adni_subpath = resolved_path / "ADNI"
        if adni_subpath.exists():
            return adni_subpath
        
        return resolved_path  # Return even if doesn't exist
    
    def _suggest_similar_folders(self, target_folder_name):
        """
        Find and suggest similar folder names.
        
        Args:
            target_folder_name (str): Target folder name to find similar ones
            
        Returns:
            list: List of suggested folder names
        """
        suggestions = []
        
        # Find similar folders in Metainformation directory
        metainfo_path = self.base_path / "Metainformation"
        if metainfo_path.exists():
            for item in metainfo_path.iterdir():
                if (item.is_dir() and 
                    target_folder_name.lower() in item.name.lower()):
                    suggestions.append(f"Metainformation/{item.name}")
        
        # Find similar folders in root
        for item in self.base_path.iterdir():
            if (item.is_dir() and 
                target_folder_name.lower() in item.name.lower()):
                suggestions.append(item.name)
        
        return suggestions

    def process_metadata_folders(self, max_files_per_type=None, 
                                selected_folders=None):
        """
        Process metadata folders and extract information.
        
        Args:
            max_files_per_type (int, optional): Maximum files per scan type
            selected_folders (list, optional): List of specific folders to process
        """
        logger.info("Starting ADNI metadata detailed analysis")
        if max_files_per_type is None:
            logger.info("File limit: Unlimited")
        else:
            logger.info(f"Maximum files per scan type: {max_files_per_type}")
        
        # Determine folders to process
        if selected_folders:
            # Process only user-specified folders
            folders_to_process = []
            invalid_folders = []
            
            for folder_input in selected_folders:
                logger.info(f"Processing folder path: {folder_input}")
                
                # Resolve path
                folder_path = self._resolve_folder_path(folder_input)
                
                # Check for XML files
                xml_folders = self._find_xml_folders(folder_path)
                
                if xml_folders:
                    folders_to_process.extend(xml_folders)
                    logger.info(f"✅ Valid folder path: {folder_path}")
                    logger.info(f"   Found {len(xml_folders)} subfolders with XML files")
                else:
                    invalid_folders.append(folder_input)
                    logger.warning(f"❌ No XML files found: {folder_path}")
            
            # Suggestions for invalid folders
            if invalid_folders:
                logger.error("Unable to process these folders:")
                for invalid_folder in invalid_folders:
                    suggestions = self._suggest_similar_folders(invalid_folder)
                    if suggestions:
                        logger.info(f"  Instead of '{invalid_folder}', try:")
                        for suggestion in suggestions[:5]:  # Max 5 suggestions
                            logger.info(f"    - {suggestion}")
                    else:
                        logger.info(f"  '{invalid_folder}': No similar folders found.")
            
            if not folders_to_process:
                logger.error("No valid folders to process!")
                return
                
        else:
            # Automatic search for all folders
            logger.info("Automatically searching all metadata folders.")
            folders_to_process = []
            
            # Search Metainformation folder
            metainfo_path = self.base_path / "Metainformation"
            if metainfo_path.exists():
                folders_to_process.extend(self._find_xml_folders(metainfo_path))
            
            # Search other metadata folders
            for folder_name in ["ADNI_PET_metadata", "ADNI_MRI_metadata"]:
                folder_path = self.base_path / folder_name
                if folder_path.exists():
                    folders_to_process.extend(self._find_xml_folders(folder_path))
            
            logger.info(f"Auto search result: {len(folders_to_process)} folders found")
        
        processed_count = 0
        
        for folder in folders_to_process:
            if not folder.exists():
                logger.warning(f"Folder does not exist: {folder}")
                continue
            
            logger.info(f"Processing folder: {folder}")
            
            # Find XML files
            xml_files = list(folder.glob("*.xml"))
            logger.info(f"XML files found: {len(xml_files)}")
            
            # Limit files by scan type
            type_counts = defaultdict(int)
            
            for xml_file in xml_files:
                # Check total file limit (None means unlimited)
                if (max_files_per_type is not None and 
                    processed_count >= max_files_per_type * 10):
                    logger.info(f"Maximum processing files reached: {processed_count}")
                    break
                
                scan_type = self.extract_scan_type(str(xml_file))
                
                # Check scan type file limit (None means unlimited)
                if (max_files_per_type is not None and 
                    type_counts[scan_type] >= max_files_per_type):
                    continue
                
                metadata = self.parse_xml_metadata(xml_file)
                
                if metadata:
                    self.scan_type_data[scan_type].append(metadata)
                    type_counts[scan_type] += 1
                    processed_count += 1
                    
                    if processed_count % 100 == 0:
                        logger.info(f"Processing completed: {processed_count} files")
        
        logger.info(f"Total processed files: {processed_count}")
        logger.info(f"Scan types found: {list(self.scan_type_data.keys())}")
        
        for scan_type, data in self.scan_type_data.items():
            logger.info(f"  {scan_type}: {len(data)} files")
    
    def create_detailed_excel(self, output_file=None):
        """
        Create detailed Excel file with scan type sheets.
        
        Args:
            output_file (str, optional): Output Excel filename
            
        Returns:
            str: Output filename or None if failed
        """
        if output_file is None:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            output_file = f"adni_detailed_metadata_{timestamp}.xlsx"
        
        logger.info(f"Creating detailed Excel file: {output_file}")
        
        if not self.scan_type_data:
            logger.error("No extracted data available!")
            return None
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Create sheet for each scan type
                for scan_type, data_list in self.scan_type_data.items():
                    if not data_list:
                        continue
                    
                    df = pd.DataFrame(data_list)
                    
                    # Sheet name length limit (Excel limitation)
                    sheet_name = scan_type[:31]
                    
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    logger.info(f"Sheet '{sheet_name}' created: {len(df)} rows")
                
                # Create summary sheet
                summary_data = []
                for scan_type, data_list in self.scan_type_data.items():
                    # Calculate average age safely
                    ages = [d['age'] for d in data_list if d.get('age')]
                    avg_age = round(sum(ages) / len(ages), 1) if ages else 'N/A'
                    
                    # Calculate male ratio
                    male_count = sum(1 for d in data_list 
                                   if d.get('gender') == 'M')
                    male_ratio = (f"{male_count / len(data_list) * 100:.1f}%" 
                                if data_list else 'N/A')
                    
                    summary_data.append({
                        'scan_type': scan_type,
                        'data_count': len(data_list),
                        'average_age': avg_age,
                        'male_ratio': male_ratio,
                        'ad_patients': sum(1 for d in data_list 
                                         if d.get('research_group') == 'AD'),
                        'mci_patients': sum(1 for d in data_list 
                                          if d.get('research_group') == 'MCI'),
                        'cn_patients': sum(1 for d in data_list 
                                         if d.get('research_group') == 'CN'),
                    })
                
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='scan_type_summary', 
                                  index=False)
                
        except Exception as e:
            logger.error(f"Excel file creation error: {e}")
            return None
        
        logger.info(f"Detailed Excel file created: {output_file}")
        return output_file


def main():
    """Main function to run the metadata extractor."""
    parser = argparse.ArgumentParser(
        description='ADNI metadata detailed extraction and Excel generation'
    )
    parser.add_argument(
        '--base-path', 
        default='.',
        help='ADNI metadata base path (default: current folder)'
    )
    parser.add_argument(
        '--output', 
        default='adni_metadata.xlsx',
        help='Output Excel filename'
    )
    parser.add_argument(
        '--max-files', 
        type=int,
        default=None,
        help='Maximum files to process per scan type (default: unlimited)'
    )
    parser.add_argument(
        '--folders',
        nargs='*',
        default=None,
        help=('Select specific folders to process '
              '(e.g., --folders ADNI1_Complete_1Yr_1.5T_metadata ADNI_PET_metadata). '
              'Process all folders if not specified')
    )
    
    args = parser.parse_args()
    
    # Run extractor
    extractor = ADNIDetailedMetadataExtractor(args.base_path)
    extractor.process_metadata_folders(
        max_files_per_type=args.max_files, 
        selected_folders=args.folders
    )
    output_file = extractor.create_detailed_excel(args.output)
    
    if output_file:
        print(f"\n✅ Detailed analysis completed! Result file: {output_file}")
        print("📊 Each scan type sheet contains organized patient metadata.")
        
        # Print processed scan types
        print("\n📋 Generated scan type sheets:")
        for scan_type, data in extractor.scan_type_data.items():
            print(f"  - {scan_type}: {len(data)} subjects")
    else:
        print("❌ Excel file creation failed.")


if __name__ == "__main__":
    main()