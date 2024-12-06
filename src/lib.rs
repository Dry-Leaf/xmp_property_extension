use std::ffi::c_void;

use windows::Win32::Foundation::*;
use windows::Win32::UI::Shell::PropertiesSystem::*;
use windows::{core::*, Win32::System::Com::*};

mod property_handler;
#[cfg(test)]
mod tests;

pub const CF_CLSID: GUID = GUID::from_u128(0x33c20ecf_3e11_42c6_9285_af2dc3cb40d8);

#[implement(IClassFactory)]
struct ClassFactory;

#[allow(non_snake_case, unused_variables)]
impl IClassFactory_Impl for ClassFactory_Impl {
    fn CreateInstance(
        &self,
        punkouter: Option<&IUnknown>,
        riid: *const GUID,
        ppvobject: *mut *mut c_void,
    ) -> Result<()> {
        if punkouter.is_some() {
            return CLASS_E_NOAGGREGATION.ok();
        }

        unsafe {
            if *riid == IInitializeWithStream::IID {
                let unknown: property_handler::PropertyHandler = Default::default();
                let ph: IInitializeWithStream = unknown.into();
                ph.query(riid,ppvobject).ok()
            } else {
                E_NOINTERFACE.ok()
            }
        }
    }

    fn LockServer(&self, flock: BOOL) -> Result<()> {
        E_NOTIMPL.ok()
    }
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

    let factory = ClassFactory {};
    let unknown: IUnknown = factory.into();

    if *rclsid == CF_CLSID {
        unknown.query(riid, pout)
    } else {
        CLASS_E_CLASSNOTAVAILABLE
    }
}
