use std::time::Duration;

use windows_volume_control::AudioController;

fn main() {
    unsafe {
        let mut controller = AudioController::init(None);
        controller.GetSessions();
        controller.GetDefaultAudioEnpointVolumeControl();
        controller.GetAllProcessSessions();
        
        // Get all available audio devices
        controller.GetAllAudioDevices();
        
        // Print all session names including application sessions
        println!("All sessions:");
        let all_sessions = controller.get_all_session_names();
        for name in all_sessions {
            println!("  - {}", name);
        }
        
        // Print only audio device names
        println!("\nAudio devices:");
        let device_names = controller.get_all_device_names();
        for name in device_names {
            println!("  - {}", name);
        }
        
        // Print output devices
        println!("\nOutput devices:");
        let output_devices = controller.get_output_device_names();
        for name in output_devices {
            println!("  - {}", name);
        }
        
        // Print input devices
        println!("\nInput devices:");
        let input_devices = controller.get_input_device_names();
        for name in input_devices {
            println!("  - {}", name);
        }
        
        // Try to get master session and control it
        let master_session = controller.get_sessions_by_name("master".to_string());
        if let Some(session) = master_session.first() {
            println!("\nMaster volume: {:.2}", session.getVolume());
        } else {
            println!("\nNo master session found");
        }
    }
}
