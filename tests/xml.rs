use windows_task_scheduler::*;

#[test]
fn test_xml() {
  let mut task = Task::default();
  task.triggers.push(Trigger::EventTrigger {
    enabled: true,
    subscription: Subscription {
      log: "System".to_string(),
      source: "Microsoft-Windows-WindowsUpdateClient".to_string(),
      event_id: None,
    },
    value_queries: vec![
      Value {
        name: "title".to_string(),
        value: r#"Event/EventData/Data[@Name='updateTitle']"#.to_string(),
      },
      Value {
        name: "id".to_string(),
        value: r#"Event/System/EventID"#.to_string(),
      },
    ],
  });

  task.actions.push(Action::Exec {
    command: r#"C:\bap.exe"#.to_string(),
    arguments: Some(r#""$(id)" "$(title)""#.to_string()),
  });

  task.settings.multiple_instances_policy = MultipleInstancesPolicy::Queue;
  task.settings.disallow_start_if_on_batteries = false;
  task.settings.stop_if_going_on_batteries = false;
  task.settings.idle_settings.stop_on_idle_end = false;
  task.settings.allow_start_on_demand = false;
  task.settings.execution_time_limit = "PT1H".to_string();

  println!("{:#?}", task);

  let xml = task.to_xml().unwrap();
  println!("{}", xml);

  // task.create_task("asdfasdf").unwrap();
}
