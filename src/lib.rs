use std::cell::RefCell;
use windows::{core::*, Win32::System::Com::*, Win32::UI::Shell::PropertiesSystem::*};

#[implement(IInitializeWithStream)] //,IPropertyStore)]
pub struct PropertyHandler {
    file_name: RefCell<String>
}

#[allow(non_snake_case, unused_variables)]
impl IInitializeWithStream_Impl for PropertyHandler_Impl {
    fn Initialize(&self, pstream: Option<&IStream>, grfmode: u32) -> Result<()> {
        let pstatstg: &mut STATSTG = &mut Default::default();
        let grfStatFlag = STATFLAG(0);

        unsafe {
            pstream.unwrap().Stat(pstatstg, grfStatFlag)?;
            let file_name = pstatstg.pwcsName.to_string();
        
            *self.file_name.borrow_mut() = file_name?.to_owned();
        }

        println!("{}", self.file_name.borrow());
         
        Ok(())
    }
}

#[cfg(test)]
mod tests;

/*
impl IPropertyStore_Impl for PropertyHandler {
    fn GetCount(&self) -> Result<u32> {
        
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> Result<()> {
        Ok(())
    }

    fn GetValue(&self, key: *const PROPERTYKEY) -> Result<PROPVARIANT> {
        
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
*/
