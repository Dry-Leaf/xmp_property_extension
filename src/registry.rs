use std::collections::HashMap;
use std::io::Result;

use crate::dll::{DEFAULT_CLSID, JXL_CLSID, MPEG_4_CLSID};

use windows::core::GUID;

use winreg::enums::*;
use winreg::RegKey;

const PROPERTY_HANDLERS_KEY: &str =
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers";

const DEF_FULLDETAILS: &str =  
    "prop:System.PropGroup.Description;System.Keywords;System.PropGroup.Image;System.Image.Dimensions;System.Image.HorizontalSize;System.Image.VerticalSize;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
const DEF_PREVIEWDETAILS: &str =
    "prop:*System.Image.Dimensions;*System.Size;*System.OfflineAvailability;*System.OfflineStatus;*System.DateCreated;*System.DateModified;*System.DateAccessed;*System.SharedWith;System.Keywords";


pub fn guid_to_string(guid: &GUID) -> String {
    format!("{{{:?}}}", guid)
}

fn register_clsid_base(module_path: &str, clsid: &windows::core::GUID) -> std::io::Result<RegKey> {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let clsid_key = hkcr.open_subkey("CLSID")?;
    let (key, _) = clsid_key.create_subkey(guid_to_string(clsid))?;
    key.set_value("", &"tag-support")?;

    let (inproc, _) = key.create_subkey("InProcServer32")?;
    inproc.set_value("", &module_path)?;
    inproc.set_value("ThreadingModel", &"Both")?;

    Ok(key)
}

pub fn register(module_path: &str) -> Result<()> {
    let clsid_map = HashMap::from([
        (&DEFAULT_CLSID, vec![".png", ".gif", ".webp"]),
        (&MPEG_4_CLSID, vec![".mp4"]),
        (&JXL_CLSID, vec![".jxl"]),
    ]);

    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

    let handlers_key = hklm.open_subkey(PROPERTY_HANDLERS_KEY)?;

    for (clsid, ext_vec) in &clsid_map {
        for ext in ext_vec {
            register_clsid_base(module_path, &clsid)?;

            let (handler_key, _) = handlers_key.create_subkey_with_flags(ext, KEY_WRITE)?;
            handler_key.set_value("", &guid_to_string(&clsid))?;

            let (system_ext_key, _) =
                hkcr.create_subkey(format!("SystemFileAssociations\\{}", ext))?;
            
            let full_details_present: Result<String> = system_ext_key.get_value("FullDetails");
            if full_details_present.is_err() {
                system_ext_key.set_value("FullDetails", &DEF_FULLDETAILS)?;
            } else {
                let full_details = full_details_present?;
                let (prop, rest) = full_details.split_at(5);
                let new_full_details = format!("{}{}{}", 
                    prop, "System.PropGroup.Description;System.Keywords;", rest);

                system_ext_key.set_value("FullDetails", &new_full_details)?;
                system_ext_key.set_value("OldFullDetails", &full_details)?;
            }
            
            let preview_details_present: Result<String> = system_ext_key.get_value("PreviewDetails");
            if preview_details_present.is_err() {
                system_ext_key.set_value("PreviewDetails", &DEF_PREVIEWDETAILS)?;
            } else {
                let preview_details = preview_details_present?;
                let new_preview_details = preview_details.clone() + ";System.Keywords"; 

                system_ext_key.set_value("PreviewDetails", &new_preview_details)?;
                system_ext_key.set_value("OldPreviewDetails", &preview_details)?;
            }
        }
    }

    Ok(())
}

pub fn unregister() -> Result<()> {
    Ok(())
}
