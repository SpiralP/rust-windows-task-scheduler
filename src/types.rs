//! https://docs.microsoft.com/en-us/windows/win32/taskschd/task-scheduler-schema

use crate::api;
use std::fmt::{self, Display, Formatter};
use xml::writer::{EmitterConfig, XmlEvent};

macro_rules! element_body {
  ($writer:expr, $name:expr, $body:expr) => {
    $writer.write(XmlEvent::start_element($name))?;
    $writer.write(XmlEvent::characters(&format!("{}", $body)))?;
    $writer.write(XmlEvent::end_element())?;
  };
}

/// default 1.2
#[derive(Debug)]
pub enum Version {
  /// Windows Vista, Windows Server 2008
  V1_2,
  /// Windows 10
  V1_4,
}

impl Default for Version {
  fn default() -> Self {
    Self::V1_2
  }
}
impl Display for Version {
  fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), fmt::Error> {
    match self {
      Self::V1_2 => write!(f, "1.2"),
      Self::V1_4 => write!(f, "1.4"),
    }
  }
}

#[derive(Debug, Default)]
pub struct Task {
  /// default 1.2
  pub version: Version,
  // pub registration_info: RegistrationInfo,
  pub triggers: Vec<Trigger>,
  // pub principals: Vec<Principal>,
  pub actions: Vec<Action>,
  pub settings: Settings,
}
impl Task {
  pub fn create_task(&self, task_name: &str) -> Result<(), Box<dyn std::error::Error>> {
    let xml = self.to_xml()?;

    api::create_task(task_name, &xml)?;

    Ok(())
  }

  pub fn to_xml(&self) -> Result<String, Box<dyn std::error::Error>> {
    let mut out = Vec::new();
    let mut writer = EmitterConfig::new()
      .perform_indent(true)
      .write_document_declaration(false)
      .create_writer(&mut out);

    writer.write(
      XmlEvent::start_element("Task")
        .attr("version", &self.version.to_string())
        .attr(
          "xmlns",
          "http://schemas.microsoft.com/windows/2004/02/mit/task",
        ),
    )?;
    {
      writer.write(XmlEvent::start_element("Triggers"))?;
      {
        for trigger in &self.triggers {
          match trigger {
            Trigger::EventTrigger {
              enabled,
              subscription,
              value_queries,
            } => {
              writer.write(XmlEvent::start_element("EventTrigger"))?;
              {
                element_body!(writer, "Enabled", enabled);
                element_body!(writer, "Subscription", subscription.to_xml()?);
                if !value_queries.is_empty() {
                  writer.write(XmlEvent::start_element("ValueQueries"))?;

                  for value in value_queries {
                    writer.write(XmlEvent::start_element("Value").attr("name", &value.name))?;
                    writer.write(XmlEvent::characters(&value.value))?;
                    writer.write(XmlEvent::end_element())?;
                  }

                  writer.write(XmlEvent::end_element())?;
                }
              }
              writer.write(XmlEvent::end_element())?;
            }
          }
        }
      } // Triggers
      writer.write(XmlEvent::end_element())?;

      writer.write(XmlEvent::start_element("Settings"))?;
      {
        element_body!(
          writer,
          "MultipleInstancesPolicy",
          self.settings.multiple_instances_policy
        );
        element_body!(
          writer,
          "DisallowStartIfOnBatteries",
          self.settings.disallow_start_if_on_batteries
        );
        element_body!(
          writer,
          "StopIfGoingOnBatteries",
          self.settings.stop_if_going_on_batteries
        );
        element_body!(
          writer,
          "AllowHardTerminate",
          self.settings.allow_hard_terminate
        );
        element_body!(
          writer,
          "StartWhenAvailable",
          self.settings.start_when_available
        );
        element_body!(
          writer,
          "RunOnlyIfNetworkAvailable",
          self.settings.run_only_if_network_available
        );

        writer.write(XmlEvent::start_element("IdleSettings"))?;
        {
          element_body!(
            writer,
            "StopOnIdleEnd",
            self.settings.idle_settings.stop_on_idle_end
          );
          element_body!(
            writer,
            "RestartOnIdle",
            self.settings.idle_settings.restart_on_idle
          );
        } // IdleSettings
        writer.write(XmlEvent::end_element())?;

        element_body!(
          writer,
          "AllowStartOnDemand",
          self.settings.allow_start_on_demand
        );
        element_body!(writer, "Enabled", self.settings.enabled);
        element_body!(writer, "Hidden", self.settings.hidden);
        element_body!(writer, "RunOnlyIfIdle", self.settings.run_only_if_idle);
        element_body!(writer, "WakeToRun", self.settings.wake_to_run);
        element_body!(
          writer,
          "ExecutionTimeLimit",
          self.settings.execution_time_limit
        );
        element_body!(writer, "Priority", self.settings.priority);
      } // Settings
      writer.write(XmlEvent::end_element())?;

      writer.write(XmlEvent::start_element("Actions"))?;
      {
        for action in &self.actions {
          match action {
            Action::Exec { command, arguments } => {
              writer.write(XmlEvent::start_element("Exec"))?;
              element_body!(writer, "Command", command);
              if let Some(arguments) = arguments {
                element_body!(writer, "Arguments", arguments);
              }
              writer.write(XmlEvent::end_element())?;
            }
          }
        }
      } // Actions
      writer.write(XmlEvent::end_element())?;
    } // Task
    writer.write(XmlEvent::end_element())?;

    Ok(String::from_utf8(out)?)
  }
}

#[derive(Debug)]
pub enum Trigger {
  EventTrigger {
    enabled: bool,
    subscription: Subscription,
    value_queries: Vec<Value>,
  },
}

#[derive(Debug)]
pub struct Subscription {
  pub log: String,
  pub source: String,
  pub event_id: Option<isize>,
}
impl Subscription {
  pub fn to_xml(&self) -> Result<String, Box<dyn std::error::Error>> {
    let mut out = Vec::new();
    let mut writer = EmitterConfig::new()
      .write_document_declaration(false)
      .create_writer(&mut out);

    writer.write(XmlEvent::start_element("QueryList"))?;
    {
      writer.write(
        XmlEvent::start_element("Query")
          .attr("Id", "0")
          .attr("Path", &self.log),
      )?;
      {
        writer.write(XmlEvent::start_element("Select").attr("Path", &self.log))?;

        if let Some(event_id) = self.event_id {
          writer.write(XmlEvent::characters(&format!(
            "*[System[Provider[@Name='{}'] and EventID={}]]",
            self.source, event_id
          )))?;
        } else {
          writer.write(XmlEvent::characters(&format!(
            "*[System[Provider[@Name='{}']]]",
            self.source
          )))?;
        }

        writer.write(XmlEvent::end_element())?;
      }
      writer.write(XmlEvent::end_element())?;
    }
    writer.write(XmlEvent::end_element())?;

    // <QueryList>
    // <Query Id="0" Path="System">
    // <Select Path="System">*[System[Provider[@Name='Microsoft-Windows-WindowsUpdateClient']]]</Select>
    // </Query>
    // </QueryList>

    Ok(String::from_utf8(out)?)
  }
}

#[derive(Debug)]
pub struct Value {
  pub name: String,
  pub value: String,
}

#[derive(Debug)]
pub enum Action {
  Exec {
    command: String,
    arguments: Option<String>,
  },
}

#[derive(Debug)]
pub struct Settings {
  /// default IgnoreNew
  pub multiple_instances_policy: MultipleInstancesPolicy,
  /// default true
  pub disallow_start_if_on_batteries: bool,
  /// default true
  pub stop_if_going_on_batteries: bool,
  /// default true
  pub allow_hard_terminate: bool,
  /// default false
  pub start_when_available: bool,
  /// default false
  pub run_only_if_network_available: bool,

  pub idle_settings: IdleSettings,

  /// default true
  pub allow_start_on_demand: bool,
  /// default true
  pub enabled: bool,
  /// default false
  pub hidden: bool,
  /// default false
  pub run_only_if_idle: bool,
  /// default false
  pub wake_to_run: bool,
  /// default PT72H
  pub execution_time_limit: String,
  /// default 7
  pub priority: u8,
}
impl Default for Settings {
  fn default() -> Self {
    Self {
      multiple_instances_policy: MultipleInstancesPolicy::default(),
      disallow_start_if_on_batteries: true,
      stop_if_going_on_batteries: true,
      allow_hard_terminate: true,
      start_when_available: false,
      run_only_if_network_available: false,
      idle_settings: IdleSettings::default(),
      allow_start_on_demand: true,
      enabled: true,
      hidden: false,
      run_only_if_idle: false,
      wake_to_run: false,
      execution_time_limit: "PT72H".to_string(),
      priority: 7,
    }
  }
}

/// default `IgnoreNew`
#[derive(Debug)]
pub enum MultipleInstancesPolicy {
  IgnoreNew,
  Queue,
}
impl Default for MultipleInstancesPolicy {
  fn default() -> Self {
    Self::IgnoreNew
  }
}
impl Display for MultipleInstancesPolicy {
  fn fmt(&self, f: &mut Formatter<'_>) -> Result<(), fmt::Error> {
    match self {
      Self::IgnoreNew => write!(f, "IgnoreNew"),
      Self::Queue => write!(f, "Queue"),
    }
  }
}

#[derive(Debug)]
pub struct IdleSettings {
  /// default true
  pub stop_on_idle_end: bool,
  /// default false
  pub restart_on_idle: bool,
}
impl Default for IdleSettings {
  fn default() -> Self {
    Self {
      stop_on_idle_end: true,
      restart_on_idle: false,
    }
  }
}
