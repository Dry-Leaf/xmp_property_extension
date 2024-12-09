#[cfg(test)]
use super::*;
use crate::property_handler::PropertyHandler;
use std::path::Path;

use windows::Win32::{
    System::Com::StructuredStorage::InitPropVariantFromStringVector,
    UI::Shell::PSGUID_SUMMARYINFORMATION,
};

use xmp_toolkit::{xmp_ns::DC, XmpMeta};

#[test]
#[allow(non_snake_case, unused_variables)]
fn xmp_test() -> Result<()> {
    let img_path = r"C:\Users\nobody\Documents\code\compiled\sample.png";
    //r"C:\Users\nobody\Pictures\eeb09e9ba001af09220cfc246b437cad.jpg";
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
    let with_tag_path = r"C:\Users\nobody\Documents\code\compiled\sample.png";
    process(with_tag_path)?;

    let without_tag_path = r"C:\Users\nobody\Documents\code\compiled\notag.png";
    process(with_tag_path)
}

#[allow(non_snake_case)]
fn process(img_path: &str) -> Result<()> {
    let middle: Vec<u16> = img_path.encode_utf16().chain(Some(0)).collect();
    let pszFile: PCWSTR = PCWSTR::from_raw(middle.as_ptr());

    let dummy_ph: PropertyHandler = Default::default();
    let ph_iu: IUnknown = dummy_ph.into();

    let cf: IClassFactory = ClassFactory.into();
    unsafe { cf.CreateInstance::<Option<&IUnknown>, IInitializeWithFile>(None)? };

    let ph: IInitializeWithFile = ph_iu.cast()?;

    unsafe { ph.Initialize(pszFile, 0)? };
    let store: IPropertyStore = ph.cast()?;

    let mut pk = PROPERTYKEY::default();

    unsafe {
        let count = store.GetCount()?;
        println!("GetCount - {:?}", count);

        for p in 0..count {
            store.GetAt(p as u32, &mut pk)?;
            println!("GetAt - {:?}", pk);

            let val = store.GetValue(&pk as *const PROPERTYKEY);
            println!("GetValue - {:?}\n", val);
        }
    }

    let caps: IPropertyStoreCapabilities = ph.cast()?;

    unsafe {
        println!(
            "IsPropertyWritable - {:?}",
            caps.IsPropertyWritable(&pk as *const PROPERTYKEY)
        );
    }

    Ok(())
}
