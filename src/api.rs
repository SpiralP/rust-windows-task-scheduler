use std::{convert::TryInto, ptr};
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

pub struct WinError {
  result: HRESULT,
  message: Option<String>,
}

use std::fmt::{self, Debug, Formatter};
impl Debug for WinError {
  fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), fmt::Error> {
    if let Some(message) = &self.message {
      write!(f, "WinError {:#X} : {}", self.result, message)
    } else {
      write!(f, "WinError {:X}", self.result)
    }
  }
}

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

pub fn create(task_name: &str, xml: &str) -> Result<(), WinError> {
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

  let p_service: &mut ITaskService = unsafe { &mut *(p_service as *mut ITaskService) };
  let _p_service = ondrop(|| {
    unsafe { p_service.Release() };
  });

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

  let mut p_root_folder: *mut ITaskFolder = ptr::null_mut();
  try_hresult!(
    unsafe { p_service.GetFolder(_bstr_t("\\"), &mut p_root_folder) },
    "Cannot get Root Folder pointer"
  );

  let p_root_folder = unsafe { &mut *p_root_folder };
  let _p_root_folder = ondrop(|| unsafe {
    p_root_folder.Release();
  });

  //  If the same task exists, remove it.
  let _ignore = unsafe { p_root_folder.DeleteTask(_bstr_t(task_name), 0) };

  //  Create the task builder object to create the task.
  let mut p_task: *mut ITaskDefinition = ptr::null_mut();
  try_hresult!(
    unsafe { p_service.NewTask(0, &mut p_task) },
    "Failed to create a task definition"
  );

  let p_task = unsafe { &mut *p_task };
  let _p_task = ondrop(|| unsafe {
    p_task.Release();
  });

  try_hresult!(unsafe { p_task.put_XmlText(_bstr_t(xml,)) });

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

  let p_registered_task = unsafe { &mut *p_registered_task };
  let _p_registered_task = ondrop(|| unsafe {
    p_registered_task.Release();
  });

  Ok(())
}

fn _bstr_t(s: &str) -> BSTR {
  let s = WideCString::from_str(s).unwrap();

  // copies input string pointer to special storage
  unsafe { SysAllocString(s.as_ptr()) }
}

#[test]
fn test_create() {
  create(
    "Useless Task",
    r#"<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
      <Triggers />
      <Settings>
        <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
        <DisallowStartIfOnBatteries>true</DisallowStartIfOnBatteries>
        <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>
        <AllowHardTerminate>true</AllowHardTerminate>
        <StartWhenAvailable>false</StartWhenAvailable>
        <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
        <IdleSettings>
          <StopOnIdleEnd>true</StopOnIdleEnd>
          <RestartOnIdle>false</RestartOnIdle>
        </IdleSettings>
        <AllowStartOnDemand>true</AllowStartOnDemand>
        <Enabled>true</Enabled>
        <Hidden>false</Hidden>
        <RunOnlyIfIdle>false</RunOnlyIfIdle>
        <WakeToRun>false</WakeToRun>
        <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
        <Priority>7</Priority>
      </Settings>
      <Actions Context="Author">
        <Exec>
          <Command>PROGRAM</Command>
        </Exec>
      </Actions>
    </Task>"#,
  )
  .unwrap();
}
