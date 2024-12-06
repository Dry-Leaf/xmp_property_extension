#[cfg(test)]
use super::*;
use crate::property_handler::PropertyHandler;
use std::ffi::c_void;
use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::{PropertiesSystem::*, SHCreateStreamOnFileEx};

#[test]
#[allow(non_snake_case, unused_variables)]
fn init_test() -> Result<()> {
    let stream: &IStream = unsafe {
        let img_path = r"C:\Users\nobody\Pictures\arc\arc38\4a640c75ee8439375004ccb05ae123df.jpg";

        let middle: Vec<u16> = img_path.encode_utf16().collect();
        let pszFile: PCWSTR = PCWSTR::from_raw(middle.as_ptr());

        &SHCreateStreamOnFileEx(pszFile, 0, 0, BOOL(0), None)?
    };

    let dummy_ph: PropertyHandler = Default::default();
    let ph_iu: IUnknown = dummy_ph.into();

    let cf: IClassFactory = ClassFactory.into();

    unsafe { cf.CreateInstance::<Option<&IUnknown>, IInitializeWithStream>(Some(&ph_iu))? };

    let ph: IInitializeWithStream = ph_iu.cast()?;

    unsafe { ph.Initialize(Some(stream), 0)? };
    let store: IPropertyStore = ph.cast()?;

    let mut pk = PROPERTYKEY::default();

    unsafe {
        println!("{:?}", store.GetCount());

        store.GetAt(0 as u32, &mut pk);
        println!("{:?}", pk);

        let val = store.GetValue(&pk as *const PROPERTYKEY);
        println!("{:?}", val);
    }

    let caps: IPropertyStoreCapabilities = ph.cast()?;

    unsafe {
        println!(
            "Writable test - {:?}",
            caps.IsPropertyWritable(&pk as *const PROPERTYKEY)
        );
    }

    Ok(())
}
