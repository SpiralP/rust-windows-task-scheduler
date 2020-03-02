use std::{
  convert::TryInto,
  error,
  fmt::{self, Debug, Display, Formatter},
  ops::Deref,
  ptr,
};
use widestring::WideCString;
use winapi::{
  ctypes::c_void,
  shared::{
    ntdef::NULL,
    rpcdce::{RPC_C_AUTHN_LEVEL_PKT, RPC_C_IMP_LEVEL_IMPERSONATE},
    winerror::HRESULT,
    wtypes::BSTR,
    wtypesbase::CLSCTX_INPROC_SERVER,
  },
  um::{
    combaseapi::{CoCreateInstance, CoInitializeEx, CoInitializeSecurity, CoUninitialize},
    oaidl::VARIANT,
    objbase::COINIT_MULTITHREADED,
    oleauto::SysAllocString,
    taskschd::{
      IRegisteredTask, ITaskDefinition, ITaskFolder, ITaskService, TaskScheduler,
      TASK_CREATE_OR_UPDATE, TASK_LOGON_INTERACTIVE_TOKEN,
    },
  },
  Class, Interface,
};

#[derive(Debug)]
pub struct WinError {
  /// https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-error-and-success-constants
  pub result: HRESULT,
  pub message: Option<String>,
}

impl Display for WinError {
  fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), fmt::Error> {
    if let Some(message) = &self.message {
      write!(f, "WinError {:#X} : {}", self.result, message)
    } else {
      write!(f, "WinError {:X}", self.result)
    }
  }
}
impl error::Error for WinError {}

macro_rules! try_hresult {
  ($expr:expr) => {
    if ::winapi::shared::winerror::FAILED($expr) {
      return Err(WinError {
        result: $expr,
        message: None,
      });
    } else {
      $expr
    }
  };
  ($expr:expr, $message:expr) => {
    if ::winapi::shared::winerror::FAILED($expr) {
      return Err(WinError {
        result: $expr,
        message: Some(format!("{}", $message)),
      });
    } else {
      $expr
    }
  };
}

struct OnDrop<T>
where
  T: FnOnce(),
{
  callback: Option<T>,
}

impl<T> Drop for OnDrop<T>
where
  T: FnOnce(),
{
  fn drop(&mut self) {
    if let Some(callback) = self.callback.take() {
      callback();
    }
  }
}

fn ondrop<T>(callback: T) -> OnDrop<T>
where
  T: FnOnce(),
{
  OnDrop {
    callback: Some(callback),
  }
}

fn with_com<T>(f: T) -> Result<(), WinError>
where
  T: FnOnce() -> Result<(), WinError>,
{
  try_hresult!(
    unsafe { CoInitializeEx(NULL, COINIT_MULTITHREADED) },
    "CoInitializeEx failed"
  );
  let _couninit = ondrop(|| unsafe {
    CoUninitialize();
  });

  //  Set general COM security levels.
  try_hresult!(
    unsafe {
      CoInitializeSecurity(
        NULL,
        -1,
        NULL as _,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL,
      )
    },
    "CoInitializeSecurity failed"
  );

  f()
}

fn with_dispatch<U, T>(dispatch: &U, f: T) -> Result<(), WinError>
where
  U: Deref<Target = winapi::um::oaidl::IDispatch>,
  T: FnOnce(&U) -> Result<(), WinError>,
{
  let _drop = ondrop(|| {
    unsafe { dispatch.Release() };
  });

  f(dispatch)
}

fn with_folder<T>(folder_path: &str, f: T) -> Result<(), WinError>
where
  T: FnOnce(&ITaskFolder, &ITaskService) -> Result<(), WinError>,
{
  with_com(|| {
    with_dispatch(
      {
        let mut p_service: *mut c_void = ptr::null_mut();
        try_hresult!(
          unsafe {
            CoCreateInstance(
              &TaskScheduler::uuidof(),
              NULL as _,
              CLSCTX_INPROC_SERVER,
              &ITaskService::uuidof(),
              &mut p_service as _,
            )
          },
          "Failed to create an instance of ITaskService"
        );

        unsafe { &*(p_service as *mut ITaskService) }
      },
      |p_service| {
        //  Connect to the task service.
        try_hresult!(
          unsafe {
            p_service.Connect(
              VARIANT::default(),
              VARIANT::default(),
              VARIANT::default(),
              VARIANT::default(),
            )
          },
          "ITaskService::Connect failed"
        );

        with_dispatch(
          {
            let mut p_root_folder: *mut ITaskFolder = ptr::null_mut();
            try_hresult!(
              unsafe { p_service.GetFolder(_bstr_t(folder_path), &mut p_root_folder) },
              "Cannot get Root Folder pointer"
            );

            unsafe { &*p_root_folder }
          },
          |p_root_folder| f(p_root_folder, p_service),
        )
      },
    )
  })
}

pub fn delete_task(task_name: &str) -> Result<(), WinError> {
  with_folder("\\", |p_root_folder, _p_service| {
    //  If the same task exists, remove it.
    try_hresult!(unsafe { p_root_folder.DeleteTask(_bstr_t(task_name), 0) });

    Ok(())
  })
}

pub fn create_task(task_name: &str, xml: &str) -> Result<(), WinError> {
  with_folder("\\", |p_root_folder, p_service| {
    //  If the same task exists, remove it.
    let _ignore = unsafe { p_root_folder.DeleteTask(_bstr_t(task_name), 0) };

    with_dispatch(
      {
        //  Create the task builder object to create the task.
        let mut p_task: *mut ITaskDefinition = ptr::null_mut();
        try_hresult!(
          unsafe { p_service.NewTask(0, &mut p_task) },
          "Failed to create a task definition"
        );

        unsafe { &*p_task }
      },
      |p_task| {
        try_hresult!(unsafe { p_task.put_XmlText(_bstr_t(xml)) });

        with_dispatch(
          {
            //  ------------------------------------------------------
            //  Save the task in the root folder.
            let mut p_registered_task: *mut IRegisteredTask = ptr::null_mut();
            try_hresult!(
              unsafe {
                p_root_folder.RegisterTaskDefinition(
                  _bstr_t(task_name),
                  p_task,
                  TASK_CREATE_OR_UPDATE.try_into().unwrap(),
                  VARIANT::default(),
                  VARIANT::default(),
                  TASK_LOGON_INTERACTIVE_TOKEN,
                  VARIANT::default(),
                  &mut p_registered_task,
                )
              },
              "Error saving the Task"
            );

            unsafe { &mut *p_registered_task }
          },
          |_p_registered_task| {
            //
            Ok(())
          },
        )
      },
    )
  })
}

fn _bstr_t(s: &str) -> BSTR {
  let s = WideCString::from_str(s).unwrap();

  // copies input string pointer to special storage
  unsafe { SysAllocString(s.as_ptr()) }
}
