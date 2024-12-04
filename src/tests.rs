#[cfg(test)]
use super::*;
use std::cell::RefCell;
use windows::Win32::{UI::Shell::SHCreateStreamOnFileEx, Foundation::*};

#[test]
#[allow(non_snake_case, unused_variables)]
fn init_test() -> Result<()> {
    let stream: &IStream = unsafe {
        let img_path = r"C:\Users\nobody\Pictures\arc\arc38\4cf20cb6d400ca27139246249567aa7f.png";

        let middle: Vec<u16> = img_path.encode_utf16().collect();

        //println!("{:?}", middle);

        let pszFile: PCWSTR = PCWSTR::from_raw(middle.as_ptr());

        &SHCreateStreamOnFileEx(pszFile,0,0,BOOL(0),None)?
    };

    let ph: IInitializeWithStream = PropertyHandler{file_name: RefCell::new(String::new()),}.into();

    unsafe {ph.Initialize(Some(stream),0);}

    Ok(())
}
