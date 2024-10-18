use process_api::get_process_info;
use session::{ApplicationSession, EndPointSession, Session};
use windows::{
    core::Interface,
    Win32::{
        Media::Audio::{
            eCapture, eMultimedia, eRender, Endpoints::IAudioEndpointVolume, IAudioSessionControl, IAudioSessionControl2, IAudioSessionEnumerator, IAudioSessionManager2, IMMDevice, IMMDeviceCollection, IMMDeviceEnumerator, ISimpleAudioVolume, MMDeviceEnumerator, DEVICE_STATE_ACTIVE
        },
        System::{
            Com::{CoCreateInstance, CoInitializeEx, CoUninitialize, CLSCTX_ALL, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED},
            ProcessStatus::K32GetProcessImageFileNameA,
            Threading::{OpenProcess, PROCESS_QUERY_INFORMATION, PROCESS_VM_READ},
        },
    },
};
use std::process::exit;
use log::error;

mod process_api;

mod session;

pub struct AudioController {
    default_device: Option<IMMDevice>,
    default_input_device: Option<IMMDevice>,
    imm_device_enumerator: Option<IMMDeviceEnumerator>,
    sessions: Vec<Box<dyn Session>>,
}

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
        // Initialize COM library
        // if let Err(err) = CoInitializeEx(Some(std::ptr::null_mut()), COINIT_MULTITHREADED) {
        //     eprintln!("ERROR: Failed to initialize COM library... {err}");
        //     return;
        // }
    
        // Get the device enumerator
        let device_enumerator: IMMDeviceEnumerator = CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_INPROC_SERVER).unwrap_or_else(|err| {
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
            let device: IMMDevice = device_collection.Item(device_index).unwrap_or_else(|err| {
                eprintln!("ERROR: Couldn't get device at index {device_index}... {err}");
                error!("ERROR: Couldn't get device at index {}... {}", device_index, err);
                exit(1);
            });
    
            let session_manager2: IAudioSessionManager2 = device.Activate(CLSCTX_INPROC_SERVER, None).unwrap_or_else(|err| {
                eprintln!("ERROR: Couldn't get AudioSessionManager for enumerating over processes... {err}");
                error!("ERROR: Couldn't get AudioSessionManager for enumerating over processes... {}", err);
                exit(1);
            });
    
            let session_enumerator: IAudioSessionEnumerator = session_manager2.GetSessionEnumerator().unwrap_or_else(|err| {
                eprintln!("ERROR: Couldn't get session enumerator... {err}");
                error!("ERROR: Couldn't get session enumerator... {}", err);
                exit(1);
            });
        
            for i in 0..session_enumerator.GetCount().unwrap() {
                let normal_session_control: Option<IAudioSessionControl> = session_enumerator.GetSession(i).ok();
                if normal_session_control.is_none() {
                    eprintln!("ERROR: Couldn't get session control of audio session...");
                    error!("ERROR: Couldn't get session control of audio session...");
                    continue;
                }
    
                let session_control: Option<IAudioSessionControl2> = normal_session_control.unwrap().cast().ok();
                if session_control.is_none() {
                    eprintln!("ERROR: Couldn't convert from normal session control to session control 2");
                    error!("ERROR: Couldn't convert from normal session control to session control 2");
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
                        eprintln!("ERROR: Couldn't get process info for pid {}", pid);
                        error!("ERROR: Couldn't get process info for pid {}", pid);
                        continue;
                    }
                };



                let audio_control: ISimpleAudioVolume = match session_control.unwrap().cast() {
                    Ok(data) => data,
                    Err(err) => {
                        eprintln!("ERROR: Couldn't get the simpleaudiovolume from session controller: {err}");
                        error!("ERROR: Couldn't get the simpleaudiovolume from session controller: {}", err);
                        continue;
                    }
                };
                // Loop through all sessions and check if the session name already exists, if it does, change name to name + 1
                let mut name = session_app_name;
                let mut counter = 2;
                while self.sessions.iter().any(|i| i.getName() == name) {
                    name = format!("{}({})", name, counter);
                    counter += 1;
                }
    
                let application_session = ApplicationSession::new(audio_control, name);
    
                self.sessions.push(Box::new(application_session));
            }
        }
    
        // Uninitialize COM
        CoUninitialize();
    }

    pub unsafe fn GetDefaultAudioEnpointVolumeControl(&mut self) {
        if self.imm_device_enumerator.is_none() {
            eprintln!("ERROR: Function called before creating enumerator");
            error!("ERROR: Function called before creating enumerator");
            return;
        }

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

        if !self.default_device.is_none() {
            let simple_audio_volume: IAudioEndpointVolume = self
                .default_device
                .clone()
                .unwrap()
                .Activate(CLSCTX_ALL, None)
                .unwrap_or_else(|err| {
                    eprintln!("ERROR: Couldn't get Endpoint volume control: {err}");
                    exit(1);
                });


            self.sessions.push(Box::new(EndPointSession::new(
                simple_audio_volume,
                "master".to_string(),
            )));    
        }

        if !self.default_input_device.is_none() {
            let simple_mic_volume: IAudioEndpointVolume = self
            .default_input_device
            .clone()
            .unwrap()
            .Activate(CLSCTX_ALL, None)
            .unwrap_or_else(|err| {
                eprintln!("ERROR: Couldn't get Endpoint volume control: {err}");
                exit(1);
            });

            self.sessions.push(Box::new(EndPointSession::new(
                simple_mic_volume,
                "mic".to_string(),
            )));
        }


    }


    pub unsafe fn get_all_session_names(&self) -> Vec<String> {
        self.sessions.iter().map(|i| i.getName()).collect()
    }

    pub unsafe fn get_session_by_name(&self, name: String) -> Option<&Box<dyn Session>> {
        self.sessions.iter().find(|i| i.getName() == name)
    }

}
