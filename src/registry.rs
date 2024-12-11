use std::collections::HashMap;
use std::io::Result;

use crate::dll::*;

use windows::core::GUID;

use winreg::enums::*;
use winreg::RegKey;

const PROPERTY_HANDLERS_KEY: &str =
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers";

const TAG_FULLDETAILS: &str =
    "prop:System.PropGroup.Description;System.Keywords;System.PropGroup.Image;System.Image.Dimensions;System.Image.HorizontalSize;System.Image.VerticalSize;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
const TAG_PREVIEWDETAILS: &str =
    "prop:*System.Image.Dimensions;*System.Size;*System.OfflineAvailability;*System.OfflineStatus;*System.DateCreated;*System.DateModified;*System.DateAccessed;*System.SharedWith;System.Keywords";

const DEF_FULLDETAILS: &str =
    "prop:System.PropGroup.Image;System.Image.Dimensions;System.Image.HorizontalSize;System.Image.VerticalSize;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
const DEF_PREVIEWDETAILS: &str =
    "prop:*System.Image.Dimensions;*System.Size;*System.OfflineAvailability;*System.OfflineStatus;*System.DateCreated;*System.DateModified;*System.DateAccessed;*System.SharedWith";

pub fn guid_to_string(guid: &GUID) -> String {
    format!("{{{:?}}}", guid)
}

fn register_clsid_base(module_path: &str, clsid: &windows::core::GUID) -> std::io::Result<RegKey> {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let clsid_key = hkcr.open_subkey("CLSID")?;
    let (key, _) = clsid_key.create_subkey(guid_to_string(clsid))?;
    key.set_value("", &"tag-support")?;
    key.set_value("DisableProcessIsolation", &(1 as u32))?;

    let (inproc, _) = key.create_subkey("InProcServer32")?;
    inproc.set_value("", &module_path)?;
    inproc.set_value("ThreadingModel", &"Both")?;

    Ok(key)
}

pub fn register(module_path: &str) -> Result<()> {
    let clsid_map = HashMap::from([
        (&MY_DEFAULT_CLSID, vec![".png", ".gif"]),
        (&MY_JXL_CLSID, vec![".jxl"]),
    ]);

    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

    let handlers_key = hklm.open_subkey(PROPERTY_HANDLERS_KEY)?;

    for (clsid, ext_vec) in &clsid_map {
        register_clsid_base(module_path, &clsid)?;
        for ext in ext_vec {
            let (handler_key, _) = handlers_key.create_subkey_with_flags(ext, KEY_WRITE)?;
            handler_key.set_value("", &guid_to_string(&clsid))?;

            let (system_ext_key, _) =
                hkcr.create_subkey(format!("SystemFileAssociations\\{}", ext))?;

            let full_details_present: Result<String> = system_ext_key.get_value("FullDetails");
            if full_details_present.is_err() {
                system_ext_key.set_value("FullDetails", &TAG_FULLDETAILS)?;
            } else {
                let full_details = full_details_present?;
                let (prop, rest) = full_details.split_at(5);
                let new_full_details = format!(
                    "{}{}{}",
                    prop, "System.PropGroup.Description;System.Keywords;", rest
                );

                system_ext_key.set_value("FullDetails", &new_full_details)?;
                system_ext_key.set_value("OldFullDetails", &full_details)?;
            }

            let preview_details_present: Result<String> =
                system_ext_key.get_value("PreviewDetails");
            if preview_details_present.is_err() {
                system_ext_key.set_value("PreviewDetails", &TAG_PREVIEWDETAILS)?;
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
    let clsid_map = HashMap::from([
        (
            &MY_DEFAULT_CLSID,
            (
                GUID::from_u128(DEFAULT_CLSID),
                vec![".png", ".gif", ".webp"],
            ),
        ),
        (&MY_JXL_CLSID, (GUID::from_u128(JXL_CLSID), vec![".jxl"])),
    ]);

    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let hklm = RegKey::predef(HKEY_LOCAL_MACHINE);

    let handlers_key = hklm.open_subkey(PROPERTY_HANDLERS_KEY)?;

    for (clsid, (orig_clsid, ext_vec)) in &clsid_map {
        print!("unreg - CLSID\\{}", &guid_to_string(&clsid));
        hkcr.delete_subkey_all(format!("CLSID\\{}", &guid_to_string(&clsid)))
            .ok();
        for ext in ext_vec {
            let (handler_key, _) = handlers_key.create_subkey_with_flags(ext, KEY_WRITE)?;
            handler_key.set_value("", &guid_to_string(&orig_clsid))?;

            let (system_ext_key, _) =
                hkcr.create_subkey(format!("SystemFileAssociations\\{}", ext))?;

            let old_full_details_present: Result<String> =
                system_ext_key.get_value("OldFullDetails");

            if !old_full_details_present.is_err() {
                let old_full_details = old_full_details_present?;
                system_ext_key.set_value("FullDetails", &old_full_details)?;
                system_ext_key.delete_value("OldFullDetails")?;
            } else {
                system_ext_key.set_value("FullDetails", &DEF_FULLDETAILS)?;
            }

            let old_preview_details_present: Result<String> =
                system_ext_key.get_value("OldPreviewDetails");

            if !old_preview_details_present.is_err() {
                let old_preview_details = old_preview_details_present?;
                system_ext_key.set_value("PreviewDetails", &old_preview_details)?;
                system_ext_key.delete_value("OldPreviewDetails")?;
            } else {
                system_ext_key.set_value("PreviewDetails", &DEF_PREVIEWDETAILS)?;
            }
        }
    }
    Ok(())
}
