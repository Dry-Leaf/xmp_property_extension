use std::cell::RefCell;
use std::path::Path;

use xmp_toolkit::{xmp_ns::DC, XmpMeta};

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::Com::{StructuredStorage::InitPropVariantFromStringVector, *},
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
    pub ext: u128,
}

#[allow(non_snake_case)]
impl IInitializeWithFile_Impl for PropertyHandler_Impl {
    fn Initialize(&self, pszfilepath: &PCWSTR, _grfmode: u32) -> Result<()> {
        //makes sure COM runtime is initialized
        #[cfg(test)]
        unsafe {
            let _ = CoIncrementMTAUsage();
        }

        log::trace!("Initialize!");
        let file_path = unsafe { pszfilepath.to_string()? };
        log::trace!("{}", file_path);

        let orig_ps: IPropertyStore = unsafe {
            let orig_clsid = GUID::from_u128(self.ext);

            //Initializing and retrieving interfaces
            let orig_init: IInitializeWithStream = CoCreateInstance(&orig_clsid, None, CLSCTX_ALL)?;
            let pstream = &SHCreateStreamOnFileEx(*pszfilepath, 0, 0, BOOL(0), None)?;

            orig_init.Initialize(pstream, STGM_READ.0)?;
            orig_init.cast()?
        };

        let orig_ps_cap: IPropertyStoreCapabilities = orig_ps.cast()?;
        *self.orig_ps.borrow_mut() = Some(orig_ps);
        *self.orig_ps_cap.borrow_mut() = Some(orig_ps_cap);

        //Reading XMP from file
        let file_path = unsafe { pszfilepath.to_string()? };
        let file_rpath = Path::new(&file_path);

        log::trace!("{}", file_path);

        println!("Getting xmp");

        let xmp_option = XmpMeta::from_file(file_rpath);
        if !xmp_option.is_err() {
            let xmp_data = xmp_option.unwrap();
            if xmp_data.contains_property(DC, "subject") {
                println!("XMP got");
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
        log::trace!("GetCount");
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();
        let gc = unsafe { ps.GetCount() }.unwrap();

        log::trace!("gc - {}", gc);

        let tag_tuple = self.tags.borrow();
        match *tag_tuple {
            Some(_) => Ok(gc + 1),
            None => Ok(gc),
        }
    }

    fn GetAt(&self, iprop: u32, pkey: *mut PROPERTYKEY) -> Result<()> {
        log::trace!("GetAt");
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();

        let tag_tuple = self.tags.borrow();

        let orig_attempt = unsafe { ps.GetAt(iprop, pkey) };
        unsafe {
            if orig_attempt.is_err() || (*pkey).pid == 0 {
                if let Some(tags) = &tag_tuple.as_ref() {
                    *pkey = PROPERTYKEY {
                        fmtid: PSGUID_SUMMARYINFORMATION,
                        pid: 5 as u32,
                    };
                    return Ok(());
                }
            }
        }

        orig_attempt
    }

    fn GetValue(&self, key: *const PROPERTYKEY) -> Result<PROPVARIANT> {
        log::trace!("GetValue");
        unsafe {
            log::trace!("key - {:?}", *key);
        }
        let binding = self.orig_ps.borrow();
        let ps = binding.as_ref().unwrap();

        let tag_tuple = self.tags.borrow();

        unsafe {
            if (*key).fmtid == PSGUID_SUMMARYINFORMATION && (*key).pid == 5 {
                if let Some(tags) = &tag_tuple.as_ref() {
                    let tag_ptrs: Vec<PCWSTR> = tags
                        .0
                        .iter()
                        .map(|t| PCWSTR::from_raw(t.as_ptr()))
                        .collect();
                    InitPropVariantFromStringVector(Some(&tag_ptrs))
                } else {
                    ps.GetValue(key)
                }
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
