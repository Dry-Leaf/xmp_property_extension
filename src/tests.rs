#[cfg(test)]
use super::*;
use crate::property_handler::PropertyHandler;
use std::ffi::OsString;
use std::os::windows::prelude::*;
use windows::Win32::UI::Shell::SHCreateStreamOnFileEx;

#[test]
#[allow(non_snake_case, unused_variables)]
fn init_test() -> Result<()> {
    let img_path = r"C:\Users\nobody\Pictures\arc\arc38\4a640c75ee8439375004ccb05ae123df.jpg";

    let middle: Vec<u16> = img_path.encode_utf16().collect();
    let pszFile: PCWSTR = PCWSTR::from_raw(middle.as_ptr());

    let stream: &IStream = unsafe { &SHCreateStreamOnFileEx(pszFile, 0, 0, BOOL(0), None)? };

    let dummy_ph: PropertyHandler = Default::default();
    let ph_iu: IUnknown = dummy_ph.into();

    let cf: IClassFactory = ClassFactory.into();
    unsafe { cf.CreateInstance::<Option<&IUnknown>, IInitializeWithFile>(None)? };

    let ph: IInitializeWithFile = ph_iu.cast()?;

    unsafe { ph.Initialize(pszFile, 0)? };
    let store: IPropertyStore = ph.cast()?;

    let mut pk = PROPERTYKEY::default();

    unsafe {
        println!("GetCount - {:?}", store.GetCount());

        store.GetAt(0 as u32, &mut pk)?;
        println!("GetAt - {:?}", pk);

        let val = store.GetValue(&pk as *const PROPERTYKEY);
        println!("GetValue - {:?}", val);
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
