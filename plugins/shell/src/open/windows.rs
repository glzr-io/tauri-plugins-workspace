use std::{ffi::OsString, os::windows::ffi::OsStrExt, path::PathBuf};

use windows::{
    core::{w, PCWSTR},
    Win32::{
        Foundation::{ERROR_FILE_NOT_FOUND, HWND},
        System::Com::CoInitialize,
        UI::{
            Shell::{
                Common::ITEMIDLIST, SHGetDesktopFolder, SHOpenFolderAndSelectItems, ShellExecuteW,
            },
            WindowsAndMessaging::SW_SHOW,
        },
    },
};

pub fn show_item_in_directory(file: PathBuf) -> crate::Result<()> {
    let dir = file
        .parent()
        .ok_or_else(|| crate::Error::NoParent(file.clone()))?;
    let dir = OsString::from(dir);
    let dir = encode_wide(dir);

    let desktop = unsafe {
        let _ = CoInitialize(None);
        SHGetDesktopFolder()?
    };

    let mut dir_item = ITEMIDLIST::default();
    unsafe {
        desktop.ParseDisplayName(
            HWND::default(),
            None,
            PCWSTR::from_raw(dir.as_ptr()),
            None,
            &mut dir_item as *mut _ as *mut _,
            0 as _,
        )?;
    }

    let file = OsString::from(file);
    let file = encode_wide(file);
    let mut file_item = ITEMIDLIST::default();
    unsafe {
        desktop.ParseDisplayName(
            HWND::default(),
            None,
            PCWSTR::from_raw(file.as_ptr()),
            None,
            &mut file_item as *mut _ as *mut _,
            0 as _,
        )?;
    }

    unsafe {
        if let Err(e) = SHOpenFolderAndSelectItems(&dir_item, Some(&[&file_item]), 0) {
            if e.code().0 == ERROR_FILE_NOT_FOUND.0 as i32 {
                ShellExecuteW(
                    HWND::default(),
                    w!("open"),
                    PCWSTR::from_raw(dir.as_ptr()),
                    PCWSTR::null(),
                    PCWSTR::null(),
                    SW_SHOW,
                );
            }
        }
    }

    Ok(())
}

fn encode_wide(string: impl AsRef<std::ffi::OsStr>) -> Vec<u16> {
    use std::iter::once;
    string.as_ref().encode_wide().chain(once(0)).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn it_works() {
        let p = PathBuf::from("D:\\amrbashir\\kal");
        show_item_in_directory(p).unwrap();
    }
}
