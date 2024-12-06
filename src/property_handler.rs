use regex::Regex;
use std::cell::RefCell;
use std::ffi::c_void;
use std::sync::LazyLock;

use windows::Win32::UI::Shell::PropertiesSystem::*;
use windows::{core::*, Win32::System::Com::*, Win32::System::Registry::*};

#[implement(IInitializeWithStream, IPropertyStore, IPropertyStoreCapabilities)]
#[derive(Default)]
pub struct PropertyHandler {
    pub orig_ps: RefCell<Option<IPropertyStore>>,
    pub orig_ps_cap: RefCell<Option<IPropertyStoreCapabilities>>,
}

static EXTENSION_REGEX: LazyLock<Regex> = LazyLock::new(|| Regex::new(r#"\.[^\.]*+$"#).unwrap());

fn get_file_type(file_name: &str) -> &str {
    let ext = EXTENSION_REGEX.find(file_name).unwrap();
    ext.as_str()
}

#[allow(non_snake_case, unused_variables)]
impl IInitializeWithStream_Impl for PropertyHandler_Impl {
    fn Initialize(&self, pstream: Option<&IStream>, grfmode: u32) -> Result<()> {
        //makes sure COM runtime is initialized
        let _ = unsafe { CoIncrementMTAUsage() };

        //Identifying file type
        let pstatstg: &mut STATSTG = &mut Default::default();
        let grfStatFlag = STATFLAG(0);

        let pstream = pstream.unwrap();
        let ext = unsafe {
            pstream.Stat(pstatstg, grfStatFlag)?;
            let file_name = pstatstg.pwcsName.to_string()?;

            get_file_type(&file_name).to_owned()
        };

        //Getting  the file type's original property handler
        let mut orig_key = String::with_capacity(85);
        orig_key.push_str(
            r#"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\"#,
        );
        orig_key.push_str(&ext);

        let inter: Vec<u16> = orig_key.encode_utf16().collect();
        let lpSubKey: PCWSTR = PCWSTR::from_raw(inter.as_ptr());

        let mut phkResult: HKEY = Default::default();

        let orig_ps: IPropertyStore = unsafe {
            //Getting GUID of original property handler
            RegOpenKeyExW(
                HKEY_LOCAL_MACHINE,
                lpSubKey,
                0,
                KEY_QUERY_VALUE,
                &mut phkResult,
            )
            .ok()?;

            let mut buffer = [0u16; 64];
            let mut size: u32 = (64 * std::mem::size_of_val(&buffer[0])) as u32;
            RegGetValueW(
                phkResult,
                None,
                None,
                RRF_RT_REG_SZ,
                None,
                Some(buffer.as_mut_ptr() as *mut c_void),
                Some(&mut size),
            )
            .ok()?;

            RegCloseKey(phkResult).ok()?;
            let orig_clsid: GUID = CLSIDFromString(PWSTR::from_raw(buffer.as_mut_ptr()))?;

            //Initializing and retrieving interfaces
            let orig_init: IInitializeWithStream = CoCreateInstance(&orig_clsid, None, CLSCTX_ALL)?;
            orig_init.Initialize(pstream, 0x00000002)?;
            orig_init.cast()?
        };

        let orig_ps_cap: IPropertyStoreCapabilities = orig_ps.cast()?;
        *self.orig_ps.borrow_mut() = Some(orig_ps);
        *self.orig_ps_cap.borrow_mut() = Some(orig_ps_cap);

        Ok(())
    }
}

#[allow(non_snake_case, unused_variables)]
impl IPropertyStore_Impl for PropertyHandler_Impl {
    fn GetCount(&self) -> Result<u32> {
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();
        unsafe { ps.GetCount() }
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> Result<()> {
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();
        unsafe { ps.GetAt(iprop, pkey) }
    }

    fn GetValue(&self, key: *const PROPERTYKEY) -> Result<PROPVARIANT> {
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();
        unsafe { ps.GetValue(key) }
    }

    fn SetValue(&self, key: *const PROPERTYKEY, propvar: *const PROPVARIANT) -> Result<()> {
        Ok(())
    }

    fn Commit(&self) -> Result<()> {
        Ok(())
    }
}

#[allow(non_snake_case, unused_variables)]
impl IPropertyStoreCapabilities_Impl for PropertyHandler_Impl {
    fn IsPropertyWritable(&self, key: *const PROPERTYKEY) -> windows::core::Result<()> {
        let binding = self.orig_ps_cap.borrow();
        let ps_cap = binding.as_ref().unwrap();
        unsafe { ps_cap.IsPropertyWritable(key) }
    }
}
