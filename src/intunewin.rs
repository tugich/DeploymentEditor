//! Native `.intunewin` packaging (replaces IntuneWinAppUtil.exe).
//!
//! Format (as produced by IntuneWinAppUtil / SvRooij.ContentPrep):
//! * the source folder is zipped (`/` separators, deflate)
//! * the zip is encrypted with AES-256-CBC / PKCS7; the encrypted file is
//!   `HMAC-SHA256 (32 bytes) | IV (16 bytes) | ciphertext`, where the HMAC covers
//!   `IV | ciphertext`
//! * `FileDigest` is the SHA-256 of the unencrypted zip
//! * the uploadable `.intunewin` file is an uncompressed zip containing
//!   `IntuneWinPackage/Contents/IntunePackage.intunewin` and
//!   `IntuneWinPackage/Metadata/Detection.xml`

use std::fs::File;
use std::io::{BufReader, BufWriter, Read, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};

use aes::cipher::{block_padding::Pkcs7, generic_array::GenericArray, BlockEncryptMut, KeyIvInit};
use anyhow::{bail, Context, Result};
use base64::Engine;
use hmac::{Hmac, Mac};
use sha2::{Digest, Sha256};

pub const ENCRYPTED_FILE_NAME: &str = "IntunePackage.intunewin";
pub const TOOL_VERSION: &str = "1.8.6.0";

/// Files of the editor itself that do not belong into the package.
const EXCLUDED_FILES: &[&str] = &["invoke-appdeploytoolkit.db", "invoke-appdeploytoolkit.ps1.template", "windows sandbox.wsb"];

#[derive(Clone, Debug, PartialEq)]
pub struct EncryptionInfo {
    pub encryption_key: String,
    pub mac_key: String,
    pub initialization_vector: String,
    pub mac: String,
    pub profile_identifier: String,
    pub file_digest: String,
    pub file_digest_algorithm: String,
}

impl EncryptionInfo {
    /// `fileEncryptionInfo` for the Graph commit request.
    pub fn to_graph(&self) -> serde_json::Value {
        serde_json::json!({
            "encryptionKey": self.encryption_key,
            "macKey": self.mac_key,
            "initializationVector": self.initialization_vector,
            "mac": self.mac,
            "profileIdentifier": self.profile_identifier,
            "fileDigest": self.file_digest,
            "fileDigestAlgorithm": self.file_digest_algorithm,
        })
    }
}

#[derive(Clone, Debug)]
pub struct EncryptedPackage {
    /// The encrypted content file (what gets uploaded to Azure Storage).
    pub path: PathBuf,
    pub unencrypted_size: u64,
    pub encrypted_size: u64,
    pub setup_file: String,
    pub info: EncryptionInfo,
    #[allow(dead_code)]
    pub file_count: usize,
}

fn b64(bytes: &[u8]) -> String {
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

fn random_bytes<const N: usize>() -> Result<[u8; N]> {
    let mut buf = [0u8; N];
    getrandom::getrandom(&mut buf).map_err(|e| anyhow::anyhow!("No secure random numbers available: {e}"))?;
    Ok(buf)
}

fn collect_files(root: &Path, dir: &Path, out: &mut Vec<PathBuf>) -> Result<()> {
    let mut entries: Vec<_> = std::fs::read_dir(dir).with_context(|| format!("Reading {}", dir.display()))?.flatten().collect();
    entries.sort_by_key(|e| e.file_name());
    for e in entries {
        let p = e.path();
        if p.is_dir() {
            collect_files(root, &p, out)?;
        } else {
            let name = e.file_name().to_string_lossy().to_ascii_lowercase();
            let top_level = p.parent() == Some(root);
            if name.ends_with(".intunewin") || (top_level && EXCLUDED_FILES.contains(&name.as_str())) {
                continue;
            }
            out.push(p);
        }
    }
    Ok(())
}

/// Zip `folder` (contents only, `/` separators) into `target`. Returns the number of files.
pub fn zip_folder(folder: &Path, target: &Path, progress: &mut dyn FnMut(usize, usize)) -> Result<usize> {
    let mut files = Vec::new();
    collect_files(folder, folder, &mut files)?;
    let out = File::create(target).with_context(|| format!("Creating {}", target.display()))?;
    let mut zip = zip::ZipWriter::new(BufWriter::new(out));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(1))
        .large_file(false);
    let total = files.len();
    for (i, f) in files.iter().enumerate() {
        let rel = f.strip_prefix(folder).unwrap_or(f);
        let name = rel.components().map(|c| c.as_os_str().to_string_lossy().to_string()).collect::<Vec<_>>().join("/");
        let size = f.metadata().map(|m| m.len()).unwrap_or(0);
        let opts = if size >= 0xFFFF_FFFF { options.large_file(true) } else { options };
        zip.start_file(name, opts)?;
        let mut src = BufReader::new(File::open(f).with_context(|| format!("Reading {} (is it open in another program?)", f.display()))?);
        std::io::copy(&mut src, &mut zip)?;
        progress(i + 1, total);
    }
    zip.finish()?.flush()?;
    Ok(total)
}

/// Encrypt `source` into `target` (IntuneWin format) and return the encryption info.
pub fn encrypt_file(source: &Path, target: &Path) -> Result<EncryptionInfo> {
    let key: [u8; 32] = random_bytes()?;
    let mac_key: [u8; 32] = random_bytes()?;
    let iv: [u8; 16] = random_bytes()?;
    encrypt_file_with(source, target, &key, &mac_key, &iv)
}

fn encrypt_file_with(source: &Path, target: &Path, key: &[u8; 32], mac_key: &[u8; 32], iv: &[u8; 16]) -> Result<EncryptionInfo> {
    const CHUNK: usize = 1 << 20; // multiple of the AES block size

    // Digest of the plain file
    let mut hasher = Sha256::new();
    {
        let mut r = BufReader::new(File::open(source)?);
        std::io::copy(&mut r, &mut hasher)?;
    }
    let digest = hasher.finalize();

    let mut input = File::open(source)?;
    let mut out = BufWriter::new(File::create(target).with_context(|| format!("Creating {}", target.display()))?);
    out.write_all(&[0u8; 32])?; // placeholder for the HMAC
    out.write_all(iv)?;

    let mut mac = <Hmac<Sha256> as Mac>::new_from_slice(mac_key).expect("HMAC accepts any key length");
    mac.update(iv);
    let mut enc = cbc::Encryptor::<aes::Aes256>::new(key.into(), iv.into());

    let mut buf = vec![0u8; CHUNK + 16];
    loop {
        // fill the buffer completely (or until EOF)
        let mut n = 0;
        while n < CHUNK {
            let r = input.read(&mut buf[n..CHUNK])?;
            if r == 0 {
                break;
            }
            n += r;
        }
        if n == CHUNK {
            for block in buf[..CHUNK].chunks_exact_mut(16) {
                enc.encrypt_block_mut(GenericArray::from_mut_slice(block));
            }
            mac.update(&buf[..CHUNK]);
            out.write_all(&buf[..CHUNK])?;
        } else {
            let ct = enc
                .encrypt_padded_mut::<Pkcs7>(&mut buf, n)
                .map_err(|_| anyhow::anyhow!("Encryption padding failed"))?;
            mac.update(ct);
            out.write_all(ct)?;
            break;
        }
    }
    let tag = mac.finalize().into_bytes();
    let mut file = out.into_inner().map_err(|e| e.into_error())?;
    file.seek(SeekFrom::Start(0))?;
    file.write_all(&tag)?;
    file.flush()?;

    Ok(EncryptionInfo {
        encryption_key: b64(key),
        mac_key: b64(mac_key),
        initialization_vector: b64(iv),
        mac: b64(&tag),
        profile_identifier: "ProfileVersion1".into(),
        file_digest: b64(&digest),
        file_digest_algorithm: "SHA256".into(),
    })
}

/// Zip + encrypt a project folder into `work_dir`. `progress(fraction, text)`.
pub fn build_encrypted(folder: &Path, setup_file: &str, work_dir: &Path, progress: &mut dyn FnMut(f32, String)) -> Result<EncryptedPackage> {
    if !folder.join(setup_file).is_file() {
        bail!("The setup file {setup_file} was not found in {}", folder.display());
    }
    std::fs::create_dir_all(work_dir)?;
    let zip_path = work_dir.join("content.zip");
    let enc_path = work_dir.join(ENCRYPTED_FILE_NAME);
    let count = zip_folder(folder, &zip_path, &mut |i, n| {
        progress(0.7 * i as f32 / n.max(1) as f32, format!("Compressing files ({i}/{n})"));
    })?;
    progress(0.75, "Encrypting package".into());
    let unencrypted_size = std::fs::metadata(&zip_path)?.len();
    let info = encrypt_file(&zip_path, &enc_path)?;
    let _ = std::fs::remove_file(&zip_path);
    let encrypted_size = std::fs::metadata(&enc_path)?.len();
    progress(1.0, "Package encrypted".into());
    Ok(EncryptedPackage {
        path: enc_path,
        unencrypted_size,
        encrypted_size,
        setup_file: setup_file.to_string(),
        info,
        file_count: count,
    })
}

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;")
}

/// `Detection.xml` in the layout written by IntuneWinAppUtil.
pub fn detection_xml(pkg: &EncryptedPackage, name: &str) -> String {
    let i = &pkg.info;
    format!(
        "<ApplicationInfo xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" ToolVersion=\"{TOOL_VERSION}\">\r\n\
  <Name>{}</Name>\r\n\
  <UnencryptedContentSize>{}</UnencryptedContentSize>\r\n\
  <FileName>{ENCRYPTED_FILE_NAME}</FileName>\r\n\
  <SetupFile>{}</SetupFile>\r\n\
  <EncryptionInfo>\r\n\
    <EncryptionKey>{}</EncryptionKey>\r\n\
    <MacKey>{}</MacKey>\r\n\
    <InitializationVector>{}</InitializationVector>\r\n\
    <Mac>{}</Mac>\r\n\
    <ProfileIdentifier>{}</ProfileIdentifier>\r\n\
    <FileDigest>{}</FileDigest>\r\n\
    <FileDigestAlgorithm>{}</FileDigestAlgorithm>\r\n\
  </EncryptionInfo>\r\n\
</ApplicationInfo>",
        xml_escape(name),
        pkg.unencrypted_size,
        xml_escape(&pkg.setup_file),
        i.encryption_key,
        i.mac_key,
        i.initialization_vector,
        i.mac,
        i.profile_identifier,
        i.file_digest,
        i.file_digest_algorithm
    )
}

/// Write the uploadable `.intunewin` file (uncompressed outer zip).
pub fn write_intunewin(pkg: &EncryptedPackage, target: &Path) -> Result<()> {
    let out = File::create(target).with_context(|| format!("Creating {}", target.display()))?;
    let mut zip = zip::ZipWriter::new(BufWriter::new(out));
    let stored = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Stored)
        .large_file(pkg.encrypted_size >= 0xFFFF_FFFF);
    zip.start_file(format!("IntuneWinPackage/Contents/{ENCRYPTED_FILE_NAME}"), stored)?;
    std::io::copy(&mut BufReader::new(File::open(&pkg.path)?), &mut zip)?;
    zip.start_file("IntuneWinPackage/Metadata/Detection.xml", zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored))?;
    zip.write_all(detection_xml(pkg, &pkg.setup_file).as_bytes())?;
    zip.finish()?.flush()?;
    Ok(())
}

/// Create `<output>/Invoke-AppDeployToolkit.intunewin` from a project folder.
pub fn create_package_file(folder: &Path, setup_file: &str, output_dir: &Path, progress: &mut dyn FnMut(f32, String)) -> Result<PathBuf> {
    let stem = Path::new(setup_file).file_stem().map(|s| s.to_string_lossy().to_string()).unwrap_or_else(|| "Package".into());
    let target = output_dir.join(format!("{stem}.intunewin"));
    if target.exists() {
        bail!("Please delete the existing package first: {}", target.display());
    }
    let work = std::env::temp_dir().join(format!("de-intunewin-{}", std::process::id()));
    let res = (|| {
        let pkg = build_encrypted(folder, setup_file, &work, progress)?;
        progress(1.0, "Writing .intunewin".into());
        write_intunewin(&pkg, &target)?;
        Ok(target.clone())
    })();
    let _ = std::fs::remove_dir_all(&work);
    res
}

#[cfg(test)]
mod tests {
    use super::*;
    use aes::cipher::BlockDecryptMut;

    fn decrypt(enc: &[u8], info: &EncryptionInfo) -> Vec<u8> {
        let d = |s: &str| base64::engine::general_purpose::STANDARD.decode(s).unwrap();
        let key = d(&info.encryption_key);
        let mac_key = d(&info.mac_key);
        let iv = &enc[32..48];
        assert_eq!(iv, &d(&info.initialization_vector)[..]);
        let mut mac = <Hmac<Sha256> as Mac>::new_from_slice(&mac_key).unwrap();
        mac.update(&enc[32..]);
        assert_eq!(&mac.finalize().into_bytes()[..], &enc[..32], "HMAC over IV+ciphertext stored first");
        assert_eq!(&enc[..32], &d(&info.mac)[..]);
        let mut buf = enc[48..].to_vec();
        let plain = cbc::Decryptor::<aes::Aes256>::new(key[..].into(), iv.into())
            .decrypt_padded_mut::<Pkcs7>(&mut buf)
            .unwrap()
            .to_vec();
        assert_eq!(&Sha256::digest(&plain)[..], &d(&info.file_digest)[..]);
        plain
    }

    #[test]
    fn encrypt_roundtrip_all_sizes() {
        let dir = std::env::temp_dir().join("de-enc-test");
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        for size in [0usize, 1, 15, 16, 17, (1 << 20) - 1, 1 << 20, (1 << 20) + 5, 3 * (1 << 20)] {
            let plain: Vec<u8> = (0..size).map(|i| (i * 31 % 251) as u8).collect();
            let src = dir.join("plain.bin");
            let dst = dir.join("enc.bin");
            std::fs::write(&src, &plain).unwrap();
            let info = encrypt_file(&src, &dst).unwrap();
            let enc = std::fs::read(&dst).unwrap();
            assert_eq!(enc.len(), 48 + (size / 16 + 1) * 16, "size {size}");
            assert_eq!(decrypt(&enc, &info), plain, "size {size}");
        }
    }

    #[test]
    fn package_file_structure() {
        let dir = std::env::temp_dir().join("de-pkg-test");
        let _ = std::fs::remove_dir_all(&dir);
        let src = dir.join("src");
        std::fs::create_dir_all(src.join("Files")).unwrap();
        std::fs::write(src.join("Invoke-AppDeployToolkit.exe"), b"MZ-setup").unwrap();
        std::fs::write(src.join("Files").join("setup.exe"), vec![7u8; 100_000]).unwrap();
        std::fs::write(src.join("Invoke-AppDeployToolkit.db"), b"editor db").unwrap();
        std::fs::write(src.join("old.intunewin"), b"x").unwrap();
        let out = dir.join("out");
        std::fs::create_dir_all(&out).unwrap();
        let path = create_package_file(&src, "Invoke-AppDeployToolkit.exe", &out, &mut |_, _| {}).unwrap();
        assert_eq!(path.file_name().unwrap(), "Invoke-AppDeployToolkit.intunewin");

        let mut outer = zip::ZipArchive::new(File::open(&path).unwrap()).unwrap();
        let names: Vec<String> = (0..outer.len()).map(|i| outer.by_index(i).unwrap().name().to_string()).collect();
        assert_eq!(names, ["IntuneWinPackage/Contents/IntunePackage.intunewin", "IntuneWinPackage/Metadata/Detection.xml"]);
        let mut xml = String::new();
        outer.by_name("IntuneWinPackage/Metadata/Detection.xml").unwrap().read_to_string(&mut xml).unwrap();
        let tag = |t: &str| xml.split(&format!("<{t}>")).nth(1).unwrap().split(&format!("</{t}>")).next().unwrap().to_string();
        assert_eq!(tag("SetupFile"), "Invoke-AppDeployToolkit.exe");
        assert_eq!(tag("FileName"), "IntunePackage.intunewin");
        let info = EncryptionInfo {
            encryption_key: tag("EncryptionKey"),
            mac_key: tag("MacKey"),
            initialization_vector: tag("InitializationVector"),
            mac: tag("Mac"),
            profile_identifier: tag("ProfileIdentifier"),
            file_digest: tag("FileDigest"),
            file_digest_algorithm: tag("FileDigestAlgorithm"),
        };
        let mut enc = Vec::new();
        outer.by_name("IntuneWinPackage/Contents/IntunePackage.intunewin").unwrap().read_to_end(&mut enc).unwrap();
        let plain = decrypt(&enc, &info);
        assert_eq!(plain.len() as u64, tag("UnencryptedContentSize").parse::<u64>().unwrap());
        let mut inner = zip::ZipArchive::new(std::io::Cursor::new(plain)).unwrap();
        let mut inner_names: Vec<String> = (0..inner.len()).map(|i| inner.by_index(i).unwrap().name().to_string()).collect();
        inner_names.sort();
        assert_eq!(inner_names, ["Files/setup.exe", "Invoke-AppDeployToolkit.exe"], "editor files excluded");
        let mut setup = Vec::new();
        inner.by_name("Files/setup.exe").unwrap().read_to_end(&mut setup).unwrap();
        assert_eq!(setup.len(), 100_000);
    }
}
