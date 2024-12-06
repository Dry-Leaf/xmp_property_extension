#[cfg(test)]
use super::*;
use crate::property_handler::PropertyHandler;
use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::{SHCreateStreamOnFileEx,PropertiesSystem::*};

#[test]
#[allow(non_snake_case, unused_variables)]
fn init_test() -> Result<()> {
    let stream: &IStream = unsafe {
        let img_path = r"C:\Users\nobody\Pictures\arc\arc38\4a640c75ee8439375004ccb05ae123df.jpg";

        let middle: Vec<u16> = img_path.encode_utf16().collect();
        let pszFile: PCWSTR = PCWSTR::from_raw(middle.as_ptr());

        &SHCreateStreamOnFileEx(pszFile,0,0,BOOL(0),None)?
    };

    let dummy_ph: PropertyHandler = Default::default();
    let ph: IInitializeWithStream = dummy_ph.into();
    unsafe {ph.Initialize(Some(stream),0)?;}

    let store: IPropertyStore = ph.cast()?;

    unsafe {
        println!("{:?}", store.GetCount());

        let mut pk = PROPERTYKEY::default();
        store.GetAt(0 as u32, &mut pk);        
        println!("{:?}", pk);

        let val = store.GetValue(&pk as *const PROPERTYKEY);
        println!("{:?}", val);
    }

    Ok(())
}
