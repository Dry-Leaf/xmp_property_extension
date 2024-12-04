//use std::cell::RefCell;
use regex::Regex;
use std::sync::LazyLock;

use windows::{core::*};
use windows::Win32::{System::Com::*, UI::Shell::PropertiesSystem::*};


#[implement(IInitializeWithStream,IPropertyStore)]
pub struct PropertyHandler;


static EXTENSION_REGEX: LazyLock<Regex> = 
    LazyLock::new(|| Regex::new(r#"\.[^\.]*+$"#).unwrap());

fn get_file_type(file_name: &str) -> &str {
    let ext = EXTENSION_REGEX.find(file_name).unwrap();
    ext.as_str()
}


#[allow(non_snake_case, unused_variables)]
impl IInitializeWithStream_Impl for PropertyHandler_Impl {
    fn Initialize(&self, pstream: Option<&IStream>, grfmode: u32) -> Result<()> {
        let pstatstg: &mut STATSTG = &mut Default::default();
        let grfStatFlag = STATFLAG(0);

        let ext = unsafe {
            pstream.unwrap().Stat(pstatstg, grfStatFlag)?;
            let file_name = pstatstg.pwcsName.to_string()?;
        
            println!("{}", file_name.clone());
            
            get_file_type(&file_name).to_owned()
        };

        println!("{}", ext);
         
        Ok(())
    }
}

#[allow(non_snake_case, unused_variables)]
impl IPropertyStore_Impl for PropertyHandler_Impl {
    fn GetCount(&self) -> Result<u32> {
        Ok(1 as u32)
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> Result<()> {
        Ok(())
    }

    fn GetValue(&self, key: *const PROPERTYKEY) -> Result<PROPVARIANT> {
        Ok(PROPVARIANT::default())
    }

    fn SetValue(
        &self,
        key: *const PROPERTYKEY,
        propvar: *const PROPVARIANT,
    ) -> Result<()> {
        Ok(())
    }

    fn Commit(&self) -> Result<()> {
        
        Ok(())
    }
}

#[cfg(test)]
mod tests;
