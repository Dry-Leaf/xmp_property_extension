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
        LibraryLoader::DisableThreadLibraryCalls,
        LibraryLoader::GetModuleFileNameW,
        SystemServices::DLL_PROCESS_ATTACH,
    },
    UI::Shell::PropertiesSystem::*,
};

pub const PNG_CLSID: GUID = GUID::from_u128(0x33c20ecf_3e11_42c6_9285_af2dc3cb40d8);
pub const JXL_CLSID: GUID = GUID::from_u128(0xee305c51_c1dd_4121_466a_117d67574bba);

static mut DLL_INSTANCE: HINSTANCE = HINSTANCE(std::ptr::null_mut());

fn get_module_path(instance: HINSTANCE) -> Result<String, HRESULT> {
    let mut path = [0u16; MAX_PATH as usize];
    let path_len = unsafe { GetModuleFileNameW(instance, &mut path) } as usize;
    String::from_utf16(&path[0..path_len]).map_err(|_| E_FAIL)
}

#[implement(IClassFactory)]
pub struct ClassFactory(pub String);

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
            if *riid == IInitializeWithFile::IID {
                let unknown: PropertyHandler = Default::default();
                let ph: IInitializeWithFile = unknown.into();
                ph.query(riid, ppvobject).ok()
            } else {
                E_NOINTERFACE.ok()
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
    if *riid != IClassFactory::IID {
        return E_UNEXPECTED;
    }

    let ext = match *rclsid {
        PNG_CLSID => ".png",
        JXL_CLSID => ".jxl",
        _ => return CLASS_E_CLASSNOTAVAILABLE,
    };

    let factory = ClassFactory(ext.to_owned());

    let unknown: IUnknown = factory.into();
    unknown.query(riid, pout)
}
