use std::collections::{HashMap, HashSet};
use std::sync::{
    atomic::{AtomicBool, Ordering},
    Arc, Mutex,
};
use std::time::{Duration, Instant};

use crate::AudioController;

const IGNORE_AFTER_TARGET_UPDATE: Duration = Duration::from_millis(250);
const IGNORE_AFTER_SET_VOLUME: Duration = Duration::from_millis(250);

#[derive(Clone)]
struct SharedState {
    polling_active: Arc<AtomicBool>,
    expected_volumes: Arc<Mutex<HashMap<String, f32>>>,
    mapped_sessions: Arc<Mutex<HashSet<String>>>,
    allow_all_expected: Arc<AtomicBool>,

    last_target_update: Arc<Mutex<HashMap<String, Instant>>>,
    last_set_volume: Arc<Mutex<HashMap<String, Instant>>>,
}

impl SharedState {
    fn is_allowed(&self, session_name: &str) -> bool {
        if self.allow_all_expected.load(Ordering::Relaxed) {
            return true;
        }
        self.mapped_sessions
            .lock()
            .unwrap()
            .contains(session_name)
    }

    fn should_ignore_restore(&self, session_name: &str) -> bool {
        let now = Instant::now();

        if let Some(t) = self
            .last_target_update
            .lock()
            .unwrap()
            .get(session_name)
            .copied()
        {
            if now.duration_since(t) < IGNORE_AFTER_TARGET_UPDATE {
                return true;
            }
        }

        let mut last_set = self.last_set_volume.lock().unwrap();
        match last_set.get(session_name).copied() {
            Some(t) if now.duration_since(t) < IGNORE_AFTER_SET_VOLUME => true,
            Some(_) => {
                last_set.remove(session_name);
                false
            }
            None => false,
        }
    }
}

pub struct VolumeMonitor {
    state: SharedState,
}

impl VolumeMonitor {
    pub fn new() -> Self {
        Self {
            state: SharedState {
                polling_active: Arc::new(AtomicBool::new(true)),
                expected_volumes: Arc::new(Mutex::new(HashMap::new())),
                mapped_sessions: Arc::new(Mutex::new(HashSet::new())),
                allow_all_expected: Arc::new(AtomicBool::new(false)),
                last_target_update: Arc::new(Mutex::new(HashMap::new())),
                last_set_volume: Arc::new(Mutex::new(HashMap::new())),
            },
        }
    }
    pub fn is_polling_active(&self) -> bool {
        self.state.polling_active.load(Ordering::Relaxed)
    }

    pub fn stop_polling(&self) {
        self.state.polling_active.store(false, Ordering::Relaxed);
    }

    pub fn set_mapped_sessions(&self, session_names: Vec<String>) {
        let allow_all = session_names.iter().any(|s| s == "unmapped");
        self.state
            .allow_all_expected
            .store(allow_all, Ordering::Relaxed);

        let mut mapped = self.state.mapped_sessions.lock().unwrap();
        mapped.clear();
        mapped.extend(session_names);
    }

    pub fn unregister_all_callbacks(
        &mut self,
        _controller: &mut AudioController,
    ) -> windows::core::Result<()> {
        Ok(())
    }

    pub fn update_expected_volume(&self, session_name: &str, volume: f32) -> windows::core::Result<()> {
        let clamped = volume.clamp(0.0, 1.0);
        self.state
            .expected_volumes
            .lock()
            .unwrap()
            .insert(session_name.to_string(), clamped);
        self.state
            .last_target_update
            .lock()
            .unwrap()
            .insert(session_name.to_string(), Instant::now());
        Ok(())
    }

    pub fn get_expected_volume(&self, session_name: &str) -> Option<f32> {
        if !self.state.is_allowed(session_name) {
            return None;
        }
        self.state
            .expected_volumes
            .lock()
            .unwrap()
            .get(session_name)
            .copied()
    }

    pub fn get_polling_session_names(&self) -> Vec<String> {
        self.state
            .expected_volumes
            .lock()
            .unwrap()
            .keys()
            .filter(|name| self.state.is_allowed(name))
            .cloned()
            .collect()
    }

    pub fn mark_setting(&self, session_name: &str) {
        self.state
            .last_set_volume
            .lock()
            .unwrap()
            .insert(session_name.to_string(), Instant::now());
    }

    pub fn unmark_setting(&self, session_name: &str) {
        self.state
            .last_set_volume
            .lock()
            .unwrap()
            .insert(session_name.to_string(), Instant::now());
    }

    pub fn should_ignore_restore(&self, session_name: &str) -> bool {
        self.state.should_ignore_restore(session_name)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ignore_after_target_update() {
        let monitor = VolumeMonitor::new();
        monitor.set_mapped_sessions(vec!["app".to_string()]);
        monitor.update_expected_volume("app", 0.5).unwrap();
        assert!(monitor.should_ignore_restore("app"));
    }

    #[test]
    fn respects_mapped_sessions_allowlist() {
        let monitor = VolumeMonitor::new();
        monitor.set_mapped_sessions(vec!["allowed".to_string()]);
        monitor.update_expected_volume("allowed", 0.5).unwrap();
        monitor.update_expected_volume("blocked", 0.2).unwrap();

        assert_eq!(monitor.get_expected_volume("allowed"), Some(0.5));
        assert_eq!(monitor.get_expected_volume("blocked"), None);
    }

    #[test]
    fn unmapped_token_allows_all_expected() {
        let monitor = VolumeMonitor::new();
        monitor.set_mapped_sessions(vec!["unmapped".to_string()]);
        monitor.update_expected_volume("anything", 0.75).unwrap();
        assert_eq!(monitor.get_expected_volume("anything"), Some(0.75));
    }
}
