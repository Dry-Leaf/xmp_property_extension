use regex::Regex;
use std::cell::RefCell;
use std::ffi::c_void;
use std::path::Path;
use std::sync::LazyLock;

use xmp_toolkit::{xmp_ns::DC, XmpMeta};

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::{
            Com::{StructuredStorage::InitPropVariantFromStringVector, *},
            Registry::*,
        },
        UI::Shell::{PropertiesSystem::*, SHCreateStreamOnFileEx, PSGUID_SUMMARYINFORMATION},
    },
};

#[derive(Clone)]
pub struct Tags(Vec<Vec<u16>>);

#[implement(IInitializeWithFile, IPropertyStore, IPropertyStoreCapabilities)]
#[derive(Default)]
pub struct PropertyHandler {
    pub orig_ps: RefCell<Option<IPropertyStore>>,
    pub orig_ps_cap: RefCell<Option<IPropertyStoreCapabilities>>,
    pub tags: RefCell<Option<Tags>>,
}

static EXTENSION_REGEX: LazyLock<Regex> = LazyLock::new(|| Regex::new(r#"\.[^\.]*+$"#).unwrap());

fn get_file_type(file_path: &str) -> &str {
    let ext = EXTENSION_REGEX.find(file_path).unwrap();
    println!("extension! - {:?}", ext);
    ext.as_str()
}

#[allow(non_snake_case)]
impl IInitializeWithFile_Impl for PropertyHandler_Impl {
    fn Initialize(&self, pszfilepath: &PCWSTR, _grfmode: u32) -> Result<()> {
        //makes sure COM runtime is initialized
        let _ = unsafe { CoIncrementMTAUsage() };

        //Identifying file type
        let file_path = unsafe { pszfilepath.to_string()? };
        let ext = get_file_type(&file_path).to_owned();

        //Getting  the file type's original property handler
        let mut orig_key = String::with_capacity(85);
        orig_key.push_str(
            r#"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\PropertySystem\\PropertyHandlers\\"#,
        );
        orig_key.push_str(&ext);

        let inter: Vec<u16> = orig_key.encode_utf16().chain(Some(0)).collect();
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
            let orig_init: IInitializeWithStream =
                CoCreateInstance(&orig_clsid, None, CLSCTX_INPROC_SERVER)?;
            let pstream = &SHCreateStreamOnFileEx(*pszfilepath, 0, 0, BOOL(0), None)?;

            orig_init.Initialize(pstream, 0x00000002)?;
            orig_init.cast()?
        };

        let orig_ps_cap: IPropertyStoreCapabilities = orig_ps.cast()?;
        *self.orig_ps.borrow_mut() = Some(orig_ps);
        *self.orig_ps_cap.borrow_mut() = Some(orig_ps_cap);

        //Reading XMP from file
        let file_rpath = Path::new(&file_path);

        let xmp_option = XmpMeta::from_file(file_rpath);
        if !xmp_option.is_err() {
            let xmp_data = xmp_option.unwrap();
            if xmp_data.contains_property(DC, "subject") {
                let tags: Vec<Vec<u16>> = xmp_data
                    .property_array(DC, "subject")
                    .map(|s| s.value.encode_utf16().chain(Some(0)).collect())
                    .collect();

                *self.tags.borrow_mut() = Some(Tags(tags));
            }
        }

        Ok(())
    }
}

#[allow(non_snake_case, unused_variables)]
impl IPropertyStore_Impl for PropertyHandler_Impl {
    fn GetCount(&self) -> Result<u32> {
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();
        let gc = unsafe { ps.GetCount() }.unwrap();

        let tag_tuple = self.tags.borrow();
        match *tag_tuple {
            Some(_) => Ok(gc + 1),
            None => Ok(gc),
        }
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> Result<()> {
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();

        let tag_tuple = self.tags.borrow();
        let iprop = match *tag_tuple {
            Some(_) => {
                if iprop == 0 {
                    unsafe {
                        *pkey = PROPERTYKEY {
                            fmtid: PSGUID_SUMMARYINFORMATION,
                            pid: 5 as u32,
                        }
                    };
                    return Ok(());
                } else {
                    iprop - 1
                }
            }
            None => iprop,
        };
        unsafe { ps.GetAt(iprop, pkey) }
    }

    fn GetValue(&self, key: *const PROPERTYKEY) -> Result<PROPVARIANT> {
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();

        let tags = self.tags.borrow();
        unsafe {
            if (*key).fmtid == PSGUID_SUMMARYINFORMATION && (*key).pid == 5 {
                let tag_ptrs: Vec<PCWSTR> = tags
                    .clone()
                    .unwrap()
                    .0
                    .iter()
                    .map(|t| PCWSTR::from_raw(t.as_ptr()))
                    .collect();
                InitPropVariantFromStringVector(Some(&tag_ptrs))
            } else {
                ps.GetValue(key)
            }
        }
    }

    fn SetValue(&self, key: *const PROPERTYKEY, propvar: *const PROPVARIANT) -> Result<()> {
        Err(windows::core::Error::new(
            WINCODEC_ERR_UNSUPPORTEDOPERATION,
            "Setter not supported",
        ))
    }

    fn Commit(&self) -> Result<()> {
        Err(windows::core::Error::new(
            WINCODEC_ERR_UNSUPPORTEDOPERATION,
            "Setter not supported",
        ))
    }
}

#[allow(non_snake_case)]
impl IPropertyStoreCapabilities_Impl for PropertyHandler_Impl {
    fn IsPropertyWritable(&self, _key: *const PROPERTYKEY) -> windows::core::Result<()> {
        Err(windows::core::Error::new(
            WINCODEC_ERR_UNSUPPORTEDOPERATION,
            "Setter not supported",
        ))
    }
}
