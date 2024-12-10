#[cfg(test)]
use std::fs;
use std::path::Path;

use crate::{dll::ClassFactory, properties::PropertyHandler};

use windows::{
    core::*,
    Win32::{
        System::Com::{StructuredStorage::InitPropVariantFromStringVector, *},
        UI::Shell::{PropertiesSystem::*, PSGUID_SUMMARYINFORMATION},
    },
};

use xmp_toolkit::{xmp_ns::DC, XmpMeta};

//#[test]
#[allow(non_snake_case, unused_variables)]
fn xmp_test() -> Result<()> {
    let img_path = r"test_images/sample.png";
    let rfile_path = Path::new(&img_path);

    let xmp_data = XmpMeta::from_file(rfile_path).unwrap();
    if !xmp_data.contains_property(DC, "subject") {
        println!("No tags.");
        return Ok(());
    }

    println!(
        "SUM INFO GUID - {:x}\n",
        PSGUID_SUMMARYINFORMATION.to_u128()
    );

    let tags: Vec<Vec<u16>> = xmp_data
        .property_array(DC, "subject")
        .map(|s| s.value.encode_utf16().chain(Some(0)).collect())
        .collect();

    let tag_ptrs: Vec<PCWSTR> = tags.iter().map(|t| PCWSTR::from_raw(t.as_ptr())).collect();

    let new_propvariant = unsafe { InitPropVariantFromStringVector(Some(&tag_ptrs)) };
    println!("{:?}", new_propvariant);

    Ok(())
}

#[test]
#[allow(non_snake_case, unused_variables)]
fn main_test() -> Result<()> {
    let test_im_dir = "test_images";

    for entry in fs::read_dir(test_im_dir)? {
        let entry = entry.unwrap().path();

        println!("{}", entry.extension().unwrap().to_str().unwrap());

        let ext = match entry.extension().unwrap().to_str().unwrap() {
            "jxl" => 0x95FFE0F8_AB15_4751_A2F3_CFAFDBF13664 as u128,
            "webm" => 0xC591F150_4106_4141_B5C1_30B2101453BD as u128,
            "mp4" => 0xf81b1b56_7613_4ee4_bc05_1fab5de5c07e as u128,
            _ => 0xA38B883C_1682_497E_97B0_0A3A9E801682 as u128,
        };

        process(entry.to_str().unwrap(), ext)?
    }

    Ok(())
}

#[allow(non_snake_case)]
fn process(img_path: &str, ext: u128) -> Result<()> {
    let middle: Vec<u16> = img_path.encode_utf16().chain(Some(0)).collect();
    let pszFile: PCWSTR = PCWSTR::from_raw(middle.as_ptr());

    let dummy_ph = PropertyHandler {
        ext,
        ..Default::default()
    };
    let ph_iu: IUnknown = dummy_ph.into();

    let ph: IInitializeWithFile = ph_iu.cast()?;

    unsafe { ph.Initialize(pszFile, 0)? };
    let store: IPropertyStore = ph.cast()?;

    let mut pk = PROPERTYKEY::default();

    unsafe {
        let count = store.GetCount()?;
        println!("GetCount - {:?}", count);

        for p in 0..count {
            store.GetAt(p as u32, &mut pk)?;
            //println!("GetAt - {:?}", pk);

            let val = store.GetValue(&pk as *const PROPERTYKEY);
            if pk.fmtid == PSGUID_SUMMARYINFORMATION {
                println!("GetAt index - {:?}\n", p);
                println!("GetValue - {:?}\n", val);
            }
        }
    }

    /*
    let caps: IPropertyStoreCapabilities = ph.cast()?;
        unsafe {
            println!(
                "IsPropertyWritable - {:?}",
                caps.IsPropertyWritable(&pk as *const PROPERTYKEY)
            );
        }
    */
    Ok(())
}
