use windows::{
    core::*, Win32::Media::Audio::Endpoints::*, Win32::Media::Audio::*, Win32::System::Com::*,
};

fn main() -> Result<()> {
    unsafe { CoInitialize(None)? };

    unsafe {
        let device_enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

        let device = device_enumerator.GetDefaultAudioEndpoint(eCapture, eConsole)?;

        let endpoint_volume: IAudioEndpointVolume = device.Activate(CLSCTX_ALL, None)?;

        let muted = endpoint_volume.GetMute()?.as_bool();

        endpoint_volume.SetMute(!muted, std::ptr::null_mut())?;

        println!("Microphone {} ", if !muted { "muted" } else { "unmuted" });
    }

    // Uninitialize COM
    unsafe { CoUninitialize() };

    Ok(())
}
