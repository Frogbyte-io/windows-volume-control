use process_api::get_process_info;
use session::{ApplicationSession, EndPointSession, Session};
use windows::{
    core::ComInterface,
    Win32::{
        Media::Audio::{
            eCapture, eMultimedia, eRender, Endpoints::IAudioEndpointVolume, IAudioSessionControl, IAudioSessionControl2, IAudioSessionEnumerator, IAudioSessionManager2, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, ISimpleAudioVolume, MMDeviceEnumerator, PKEY_AudioEndpoint_FormFactor, DEVICE_STATE_ACTIVE
        },
        System::{
            Com::{CoCreateInstance, CoInitializeEx, CLSCTX_ALL, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED, STGM_READ, CoTaskMemFree},
        },
        Devices::FunctionDiscovery::PKEY_Device_FriendlyName,
        UI::Shell::PropertiesSystem::{IPropertyStore, PropVariantToStringAlloc},
    },
};
use std::process::exit;
use log::error;

mod process_api;

mod session;
mod volume_monitor;

const FORM_FACTOR_SPDIF: i32 = 8;

pub use volume_monitor::VolumeMonitor;

// Helper function to get device friendly name using PropVariantToStringAlloc
fn get_device_friendly_name(device: &IMMDevice, fallback_name: &str) -> String {
    unsafe {
        // Start with the caller-supplied fallback text. We will overwrite it only
        // if we can obtain a readable property value.
        let mut friendly_name = fallback_name.to_string();
        if let Ok(property_store) = device.OpenPropertyStore(STGM_READ) {
            if let Ok(prop_variant) = property_store.GetValue(&PKEY_Device_FriendlyName) {
                // Use PropVariantToStringAlloc which handles conversion and allocation
                match PropVariantToStringAlloc(&prop_variant) {
                    Ok(buffer) => {
                        if !buffer.is_null() {
                            if let Ok(name) = buffer.to_string() {
                                if !name.is_empty() {
                                    friendly_name = name; // Success!
                                }
                            }
                            // Free the buffer allocated by PropVariantToStringAlloc
                            CoTaskMemFree(Some(buffer.as_ptr().cast()));
                        }
                    }
                    Err(_err) => {
                        // Log error if needed: eprintln!("PropVariantToStringAlloc failed: {:?}", e);
                        // Keep the fallback name if conversion fails
                    }
                }
                // No need for manual PropVariantClear when using PropVariantToStringAlloc typically,
                // and windows-rs PROPVARIANT might implement Drop correctly.
            }
        }
        friendly_name
    }
}

fn get_device_form_factor(property_store: &IPropertyStore) -> Option<i32> {
    unsafe {
        if let Ok(prop_variant) = property_store.GetValue(&PKEY_AudioEndpoint_FormFactor) {
            if let Ok(buffer) = PropVariantToStringAlloc(&prop_variant) {
                if !buffer.is_null() {
                    let form_factor = buffer
                        .to_string()
                        .ok()
                        .and_then(|val| val.parse::<i32>().ok());
                    CoTaskMemFree(Some(buffer.as_ptr().cast()));
                    return form_factor;
                }
            }
        }
        None
    }
}

pub struct AudioController {
    default_device: Option<IMMDevice>,
    default_input_device: Option<IMMDevice>,
    imm_device_enumerator: Option<IMMDeviceEnumerator>,
    sessions: Vec<Box<dyn Session>>,
    default_output_id: Option<String>,
    default_input_id: Option<String>,
}

#[derive(Debug)]
pub enum CoinitMode {
    MultiTreaded,
    ApartmentThreaded
}

impl AudioController {
    pub unsafe fn init(coinit_mode: Option<CoinitMode>) -> Self {
        let mut coinit: windows::Win32::System::Com::COINIT = COINIT_MULTITHREADED;
        if let Some(x) = coinit_mode {
            match x {
                CoinitMode::ApartmentThreaded   => {coinit = COINIT_APARTMENTTHREADED},
                CoinitMode::MultiTreaded        => {coinit = COINIT_MULTITHREADED}
            }
        }

        CoInitializeEx(None, coinit).unwrap_or_else(|err| {
            eprintln!("ERROR: Couldn't initialize windows connection: {err}");
            error!("ERROR: Couldn't initialize windows connection: {}", err);
            exit(1);
    });

        Self {
            default_device: None,
            default_input_device: None,
            imm_device_enumerator: None,
            sessions: vec![],
            default_output_id: None,
            default_input_id: None,
        }
    }

    pub unsafe fn GetSessions(&mut self) {
        self.imm_device_enumerator = Some(
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER).unwrap_or_else(
                |err| {
                    eprintln!("ERROR: Couldn't get Media device enumerator: {err}");
                    error!("ERROR: Couldn't get Media device enumerator: {}", err);
                    exit(1);
                },
            ),
        );
    }


    pub unsafe fn GetAllProcessSessions(&mut self) {
        // Get the device enumerator
        let device_enumerator_result = CoCreateInstance::<_, IMMDeviceEnumerator>(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER);
        let device_enumerator: IMMDeviceEnumerator = device_enumerator_result.unwrap_or_else(|err| {
            eprintln!("ERROR: Couldn't create device enumerator... {err}");
            error!("ERROR: Couldn't create device enumerator... {}", err);
            exit(1);
        });
    
        // Get all audio output devices
        let device_collection: IMMDeviceCollection = device_enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE).unwrap_or_else(|err| {
            eprintln!("ERROR: Couldn't enumerate audio endpoints... {err}");
            error!("ERROR: Couldn't enumerate audio endpoints... {}", err);
            exit(1);
        });
    
        let device_count = device_collection.GetCount().unwrap_or_else(|err| {
            eprintln!("ERROR: Couldn't get device count... {err}");
            error!("ERROR: Couldn't get device count... {}", err);
            exit(1);
        });
    
        for device_index in 0..device_count {
            let device: IMMDevice = match device_collection.Item(device_index) {
                Ok(dev) => dev,
                Err(err) => {
                    eprintln!("WARNING: Skipping device {} - couldn't get device: {}", device_index, err);
                    error!("WARNING: Skipping device {} - couldn't get device: {}", device_index, err);
                    continue;
                }
            };

            // Get device ID for error reporting
            let device_id = match device.GetId() {
                Ok(id) => match id.to_string() {
                    Ok(id_str) => id_str,
                    Err(_) => {
                        eprintln!("WARNING: Skipping device {} - couldn't convert device ID to string", device_index);
                        error!("WARNING: Skipping device {} - couldn't convert device ID to string", device_index);
                        continue;
                    }
                },
                Err(err) => {
                    eprintln!("WARNING: Skipping device {} - couldn't get device ID: {}", device_index, err);
                    error!("WARNING: Skipping device {} - couldn't get device ID: {}", device_index, err);
                    continue;
                }
            };

            // Attempt to open property store - if this fails, do not proceed
            let property_store = match device.OpenPropertyStore(STGM_READ) {
                Ok(store) => store,
                Err(err) => {
                    let device_name = get_device_friendly_name(&device, &format!("Device {}", device_index));
                    eprintln!("ERROR: Skipping device '{}' (ID: {}) - couldn't open property store: {}", device_name, device_id, err);
                    error!("ERROR: Skipping device '{}' (ID: {}) - couldn't open property store: {}", device_name, device_id, err);
                    continue;
                }
            };

            let device_name = get_device_friendly_name(&device, &format!("Device {}", device_index));

            // Skip SPDIF/digital outputs if detected by form factor
            if let Some(form_factor) = get_device_form_factor(&property_store) {
                if form_factor == FORM_FACTOR_SPDIF {
                    continue;
                }
            }
    
            let session_manager2: IAudioSessionManager2 = match device.Activate(CLSCTX_INPROC_SERVER, None) {
                Ok(mgr) => mgr,
                Err(err) => {
                    eprintln!("WARNING: Skipping device '{}' (ID: {}) - couldn't activate AudioSessionManager: {}", device_name, device_id, err);
                    error!("WARNING: Skipping device '{}' (ID: {}) - couldn't activate AudioSessionManager: {}", device_name, device_id, err);
                    continue;
                }
            };
    
            let session_enum_result = session_manager2.GetSessionEnumerator();
            // If GetSessionEnumerator fails for this device, skip it
            let session_enumerator: IAudioSessionEnumerator = match session_enum_result {
                Ok(enumerator) => enumerator,
                Err(err) => {
                    eprintln!("WARNING: Skipping device '{}' (ID: {}) - couldn't get session enumerator: {}", device_name, device_id, err);
                    error!("WARNING: Skipping device '{}' (ID: {}) - couldn't get session enumerator: {}", device_name, device_id, err);
                    continue;
                }
            };
        
            let session_count = match session_enumerator.GetCount() {
                Ok(count) => count,
                Err(err) => {
                    eprintln!("WARNING: Skipping device '{}' (ID: {}) - couldn't get session count: {}", device_name, device_id, err);
                    error!("WARNING: Skipping device '{}' (ID: {}) - couldn't get session count: {}", device_name, device_id, err);
                    continue;
                }
            };
            
            for i in 0..session_count {
                let normal_session_control: Option<IAudioSessionControl> = session_enumerator.GetSession(i).ok();
                if normal_session_control.is_none() {
                    eprintln!("ERROR: Couldn't get session control of audio session for device '{}' (ID: {})", device_name, device_id);
                    error!("ERROR: Couldn't get session control of audio session for device '{}' (ID: {})", device_name, device_id);
                    continue;
                }
    
                let session_control: Option<IAudioSessionControl2> = normal_session_control.unwrap().cast().ok();
                if session_control.is_none() {
                    eprintln!("ERROR: Couldn't convert from normal session control to session control 2 for device '{}' (ID: {})", device_name, device_id);
                    error!("ERROR: Couldn't convert from normal session control to session control 2 for device '{}' (ID: {})", device_name, device_id);
                    continue;
                }
    
                let pid = session_control.as_ref().unwrap().GetProcessId().unwrap();
                if pid == 0 {
                    continue;
                }

               let session_app_name = match get_process_info(pid) {
                    Ok(info) => {
                        info.process_name.clone()
                    },
                    Err(_err) => {
                        eprintln!("ERROR: Couldn't get process info for pid {} on device '{}' (ID: {})", pid, device_name, device_id);
                        error!("ERROR: Couldn't get process info for pid {} on device '{}' (ID: {})", pid, device_name, device_id);
                        continue;
                    }
                };

                let audio_control: ISimpleAudioVolume = match session_control.unwrap().cast() {
                    Ok(data) => data,
                    Err(err) => {
                        eprintln!("ERROR: Couldn't get the simpleaudiovolume from session controller for device '{}' (ID: {}): {}", device_name, device_id, err);
                        error!("ERROR: Couldn't get the simpleaudiovolume from session controller for device '{}' (ID: {}): {}", device_name, device_id, err);
                        continue;
                    }
                };

                let name = session_app_name;
    
                let application_session = ApplicationSession::new(audio_control, name);
    
                self.sessions.push(Box::new(application_session));
            }
        }
    
    }

    pub unsafe fn GetDefaultAudioEnpointVolumeControl(&mut self) {
        if self.imm_device_enumerator.is_none() {
            eprintln!("ERROR: Function called before creating enumerator");
            error!("ERROR: Function called before creating enumerator");
            return;
        }

        // Get Default Output Device
        self.default_device = match self.imm_device_enumerator
            .clone()
            .unwrap()
            .GetDefaultAudioEndpoint(eRender, eMultimedia)
            {
                Ok(device) => Some(device),
                Err(err) => {
                    eprintln!("ERROR: Couldn't get Default audio output endpoint {err}");
                    None
                }
            };

        // Get Default Input Device
        self.default_input_device = match self.imm_device_enumerator
            .clone()
            .unwrap()
            .GetDefaultAudioEndpoint(eCapture, eMultimedia)
            {
                Ok(device) => Some(device),
                Err(err) => {
                    eprintln!("ERROR: Couldn't get Default audio input endpoint {err}");
                    None
                }
            };

        // Process Default Output Device
        if let Some(ref device) = self.default_device {
            let device_id_str = device.GetId().ok().and_then(|id| id.to_string().ok());
            self.default_output_id = device_id_str.clone(); // Store the ID

            let endpoint_volume: IAudioEndpointVolume = match device.Activate(CLSCTX_ALL, None) {
                Ok(volume) => volume,
                Err(err) => {
                    eprintln!("ERROR: Couldn't activate Endpoint volume control for default output: {err}");
                    error!("ERROR: Couldn't activate Endpoint volume control for default output: {}", err);
                    return; // Exit if activation fails
                }
            };

            let friendly_name = get_device_friendly_name(device, "Default Output");
            let session_name = format!("output: {} (default_output)", friendly_name);
            
            self.sessions.push(Box::new(EndPointSession::new(
                endpoint_volume,
                session_name,
            )));    
            
            // For backwards compatibility, also add a new controller with the old "master" name
            if let Ok(master_endpoint_volume) = device.Activate(CLSCTX_ALL, None) {
                self.sessions.push(Box::new(EndPointSession::new(
                    master_endpoint_volume,
                    "master".to_string(),
                )));
            }
        }

        // Process Default Input Device
        if let Some(ref device) = self.default_input_device {
            let device_id_str = device.GetId().ok().and_then(|id| id.to_string().ok());
            self.default_input_id = device_id_str.clone(); // Store the ID

            let endpoint_volume: IAudioEndpointVolume = match device.Activate(CLSCTX_ALL, None) {
                Ok(volume) => volume,
                Err(err) => {
                    eprintln!("ERROR: Couldn't activate Endpoint volume control for default input: {err}");
                    error!("ERROR: Couldn't activate Endpoint volume control for default input: {}", err);
                    return; // Exit if activation fails
                }
            };

            let friendly_name = get_device_friendly_name(device, "Default Input");
            let session_name = format!("input: {} (default_input)", friendly_name);

            self.sessions.push(Box::new(EndPointSession::new(
                endpoint_volume,
                session_name,
            )));
            
            // For backwards compatibility, also add a new controller with the old "mic" name
            if let Ok(mic_endpoint_volume) = device.Activate(CLSCTX_ALL, None) {
                self.sessions.push(Box::new(EndPointSession::new(
                    mic_endpoint_volume,
                    "mic".to_string(),
                )));
            }
        }
    }

    pub unsafe fn GetAllAudioDevices(&mut self) {
        if self.imm_device_enumerator.is_none() {
            self.imm_device_enumerator = Some(
                CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER).unwrap_or_else(
                    |err| {
                        eprintln!("ERROR: Couldn't get Media device enumerator: {err}");
                        error!("ERROR: Couldn't get Media device enumerator: {}", err);
                        exit(1);
                    },
                ),
            );
        }

        // Get all output (render) devices
        let output_device_collection: IMMDeviceCollection = match
            self.imm_device_enumerator
                .as_ref()
                .unwrap()
                .EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)
        {
            Ok(col) => col,
            Err(err) => {
                eprintln!("ERROR: Couldn't enumerate output endpoints: {err}");
                error!("ERROR: Couldn't enumerate output endpoints: {}", err);
                return;
            }
        };
        let output_device_count = output_device_collection.GetCount().unwrap();

        for device_index in 0..output_device_count {
            if let Ok(device) = output_device_collection.Item(device_index) {
                // Check if this device is the default output device
                let is_default = if let Some(ref default_id) = self.default_output_id {
                    device.GetId().ok().and_then(|id| id.to_string().ok()) == Some(default_id.clone())
                } else {
                    false
                };

                if is_default {
                    continue; // Skip default device, already added
                }

                // Get device friendly name
                let friendly_name = get_device_friendly_name(&device, "Unknown Output Device");

                // Create endpoint volume controller for this device
                if let Ok(endpoint_volume) = device.Activate::<IAudioEndpointVolume>(CLSCTX_ALL, None) {
                    self.sessions.push(Box::new(EndPointSession::new(
                        endpoint_volume,
                        format!("output: {}", friendly_name),
                    )));
                } else {
                     eprintln!("ERROR: Couldn't activate Endpoint volume for output device: {}", friendly_name);
                     error!("ERROR: Couldn't activate Endpoint volume for output device: {}", friendly_name);
                }
            }
        }

        // Get all input (capture) devices
        let input_device_collection: IMMDeviceCollection = match self.imm_device_enumerator.as_ref().unwrap().EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE) {
            Ok(col) => col,
            Err(err) => {
                eprintln!("ERROR: Couldn't enumerate input endpoints: {err}");
                error!("ERROR: Couldn't enumerate input endpoints: {}", err);
                return;
            }
        };
        let input_device_count = match input_device_collection.GetCount() {
            Ok(count) => count,
            Err(err) => {
                eprintln!("ERROR: Couldn't get input device count: {err}");
                error!("ERROR: Couldn't get input device count: {}", err);
                return;
            }
        };

        for device_index in 0..input_device_count {
             if let Ok(device) = input_device_collection.Item(device_index) {
                // Check if this device is the default input device
                let is_default = if let Some(ref default_id) = self.default_input_id {
                    device.GetId().ok().and_then(|id| id.to_string().ok()) == Some(default_id.clone())
                } else {
                    false
                };

                if is_default {
                    continue; // Skip default device, already added
                }

                // Get device friendly name
                let friendly_name = get_device_friendly_name(&device, "Unknown Input Device");

                // Create endpoint volume controller for this device
                if let Ok(endpoint_volume) = device.Activate::<IAudioEndpointVolume>(CLSCTX_ALL, None) {
                    self.sessions.push(Box::new(EndPointSession::new(
                        endpoint_volume,
                        format!("input: {}", friendly_name),
                    )));
                } else {
                     eprintln!("ERROR: Couldn't activate Endpoint volume for input device: {}", friendly_name);
                     error!("ERROR: Couldn't activate Endpoint volume for input device: {}", friendly_name);
                }
            }
        }
    }

    //returns all session names
    pub unsafe fn get_all_session_names(&self) -> Vec<String> {
        self.sessions.iter().map(|i| i.getName()).collect()
    }

    //returns all audio device names (both input and output)
    pub unsafe fn get_all_device_names(&self) -> Vec<String> {
        self.sessions.iter()
            .map(|i| i.getName())
            .filter(|name| 
                name.starts_with("input: ") || 
                name.starts_with("output: ")
            )
            .collect()
    }

    //returns all audio output device names
    pub unsafe fn get_output_device_names(&self) -> Vec<String> {
        self.sessions.iter()
            .map(|i| i.getName())
            .filter(|name| 
                name.starts_with("output: ")
            )
            .collect()
    }

    //returns all audio input device names
    pub unsafe fn get_input_device_names(&self) -> Vec<String> {
        self.sessions.iter()
            .map(|i| i.getName())
            .filter(|name| 
                name.starts_with("input: ")
            )
            .collect()
    }

    //returns all sessions with the given name
    pub unsafe fn get_sessions_by_name(&self, name: String) -> Vec<&Box<dyn Session>> {
        self.sessions.iter().filter(|i| i.getName() == name).collect()
    }

    //returns the first session with the given name
    pub unsafe fn get_session_by_name(&self, name: String) -> Option<&Box<dyn Session>> {
        self.sessions.iter().find(|i| i.getName() == name)
    }
}
