use std::ffi::c_void;

use crate::{
    properties::PropertyHandler,
    registry::{register, unregister},
};

use windows::core::{implement, IUnknown, Interface, GUID, HRESULT};
use windows::Win32::{
    Foundation::*,
    System::{
        Com::{IClassFactory, IClassFactory_Impl},
        LibraryLoader::{DisableThreadLibraryCalls, GetModuleFileNameW},
        SystemServices::DLL_PROCESS_ATTACH,
    },
    UI::Shell::PropertiesSystem::*,
};

pub const DEFAULT_CLSID: u128 = 0xA38B883C_1682_497E_97B0_0A3A9E801682;
pub const JXL_CLSID: u128 = 0x95FFE0F8_AB15_4751_A2F3_CFAFDBF13664;

pub const MY_DEFAULT_CLSID: GUID = GUID::from_u128(0x33C20ECF_3E11_42C6_9285_AF2DC3CB40D8);
pub const MY_JXL_CLSID: GUID = GUID::from_u128(0xEE305C51_C1DD_4121_466A_117D67574BBA);

static mut DLL_INSTANCE: HINSTANCE = HINSTANCE(std::ptr::null_mut());

fn get_module_path(instance: HINSTANCE) -> Result<String, HRESULT> {
    let mut path = [0u16; MAX_PATH as usize];
    let path_len = unsafe { GetModuleFileNameW(instance, &mut path) } as usize;
    String::from_utf16(&path[0..path_len]).map_err(|_| E_FAIL)
}

#[implement(IClassFactory)]
pub struct ClassFactory(pub u128);

#[allow(non_snake_case, unused_variables)]
impl IClassFactory_Impl for ClassFactory_Impl {
    fn CreateInstance(
        &self,
        punkouter: Option<&IUnknown>,
        riid: *const GUID,
        ppvobject: *mut *mut c_void,
    ) -> windows::core::Result<()> {
        if punkouter.is_some() {
            return CLASS_E_NOAGGREGATION.ok();
        }

        unsafe {
            match *riid {
                IPropertyStore::IID | IInitializeWithStream::IID => {
                    let unknown: IUnknown = PropertyHandler {
                        ext: self.0.clone(),
                        ..Default::default()
                    }
                    .into();
                    unknown.query(riid, ppvobject).ok()
                }
                _ => {
                    log::trace!("Unknown IID: {:?}", *riid);
                    E_NOINTERFACE.ok()
                }
            }
        }
    }

    fn LockServer(&self, flock: BOOL) -> windows::core::Result<()> {
        E_NOTIMPL.ok()
    }
}

fn shell_change_notify() {
    use windows::Win32::UI::Shell::{SHChangeNotify, SHCNE_ASSOCCHANGED, SHCNF_IDLIST};
    unsafe { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None) };
}

#[no_mangle]
#[allow(non_snake_case)]
pub unsafe extern "system" fn DllRegisterServer() -> HRESULT {
    let module_path = match get_module_path(DLL_INSTANCE) {
        Ok(path) => path,
        Err(err) => return err,
    };
    if register(&module_path).is_ok() {
        shell_change_notify();
        S_OK
    } else {
        E_FAIL
    }
}

#[no_mangle]
#[allow(non_snake_case)]
pub unsafe extern "system" fn DllUnregisterServer() -> HRESULT {
    if unregister().is_ok() {
        shell_change_notify();
        S_OK
    } else {
        E_FAIL
    }
}

#[no_mangle]
#[allow(non_snake_case)]
pub extern "stdcall" fn DllMain(
    dll_instance: HINSTANCE,
    reason: u32,
    _reserved: *mut c_void,
) -> bool {
    if reason == DLL_PROCESS_ATTACH {
        unsafe {
            DLL_INSTANCE = dll_instance;
            DisableThreadLibraryCalls(dll_instance).unwrap();
        }
    }
    true
}

#[no_mangle]
#[allow(non_snake_case, unused_variables)]
pub unsafe extern "system" fn DllGetClassObject(
    rclsid: *const GUID,
    riid: *const GUID,
    pout: *mut *mut core::ffi::c_void,
) -> HRESULT {
    #[cfg(debug_assertions)]
    {
        // Set up logging to the project directory.
        simple_logging::log_to_file(
            format!("{}\\debug.log", env!("CARGO_MANIFEST_DIR")),
            log::LevelFilter::Trace,
        )
        .unwrap();
    }
    log::trace!("DllGetClassObject");
    if *riid != IClassFactory::IID {
        return E_UNEXPECTED;
    }

    let ext = match *rclsid {
        MY_DEFAULT_CLSID => DEFAULT_CLSID,
        MY_JXL_CLSID => JXL_CLSID,
        _ => DEFAULT_CLSID,
    };

    let factory = ClassFactory(ext.to_owned());

    let unknown: IUnknown = factory.into();
    unknown.query(riid, pout)
}
