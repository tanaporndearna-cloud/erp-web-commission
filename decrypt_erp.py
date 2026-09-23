"""
Decrypt an ECMA-376 Agile-encrypted xlsx (Office 365 password protection).

Usage:
    python3 decrypt_erp.py <input_encrypted.xlsx> <output_decrypted.xlsx> [password]

Default password "weTeam15" is used if not supplied.
If the input file is NOT encrypted (plain xlsx / not a CFB container),
this script just copies it to the output path unchanged.
"""
import struct, hashlib, sys, shutil
import xml.etree.ElementTree as ET
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

IN_PATH = sys.argv[1]
OUT_PATH = sys.argv[2]
PASSWORD = sys.argv[3] if len(sys.argv) > 3 else 'weTeam15'


class CFB:
    def __init__(self, data):
        self.data = data
        hdr = data[:512]
        if hdr[:8] != b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            raise ValueError('not a CFB container')
        self.sector_shift = struct.unpack('<H', hdr[30:32])[0]
        self.mini_sector_shift = struct.unpack('<H', hdr[32:34])[0]
        self.sector_size = 1 << self.sector_shift
        self.mini_sector_size = 1 << self.mini_sector_shift
        self.num_fat_sectors = struct.unpack('<I', hdr[44:48])[0]
        self.first_dir_sector = struct.unpack('<I', hdr[48:52])[0]
        self.mini_cutoff = struct.unpack('<I', hdr[56:60])[0]
        self.first_minifat_sector = struct.unpack('<I', hdr[60:64])[0]
        self.num_minifat_sectors = struct.unpack('<I', hdr[64:68])[0]
        self.first_difat_sector = struct.unpack('<I', hdr[68:72])[0]
        self.num_difat_sectors = struct.unpack('<I', hdr[72:76])[0]
        difat_header = list(struct.unpack('<109I', hdr[76:76+436]))
        difat = [x for x in difat_header if x != 0xFFFFFFFF]
        sect = self.first_difat_sector
        for _ in range(self.num_difat_sectors):
            s = self._sector_data(sect)
            entries = list(struct.unpack('<128I', s))
            sect = entries[-1]
            for x in entries[:-1]:
                if x != 0xFFFFFFFF:
                    difat.append(x)
        self.fat = []
        for fsect in difat:
            s = self._sector_data(fsect)
            n = self.sector_size // 4
            self.fat.extend(struct.unpack(f'<{n}I', s))
        dir_chain = self._chain(self.first_dir_sector)
        dir_bytes = b''.join(self._sector_data(s) for s in dir_chain)
        self.entries = []
        for i in range(0, len(dir_bytes), 128):
            e = dir_bytes[i:i+128]
            namelen = struct.unpack('<H', e[64:66])[0]
            name = e[0:max(namelen-2,0)].decode('utf-16-le') if namelen>0 else ''
            obj_type = e[66]
            start_sector = struct.unpack('<I', e[116:120])[0]
            size = struct.unpack('<Q', e[120:128])[0]
            self.entries.append({'name': name, 'type': obj_type, 'start': start_sector, 'size': size})
        self.minifat = []
        if self.first_minifat_sector != 0xFFFFFFFE:
            mf_chain = self._chain(self.first_minifat_sector)
            mf_bytes = b''.join(self._sector_data(s) for s in mf_chain)
            n = len(mf_bytes)//4
            self.minifat = list(struct.unpack(f'<{n}I', mf_bytes))
        root = [e for e in self.entries if e['type']==5][0]
        self.root = root
        root_chain = self._chain(root['start'])
        self.ministream = b''.join(self._sector_data(s) for s in root_chain)[:root['size']]

    def _sector_data(self, idx):
        off = self.sector_size * (idx+1)
        return self.data[off:off+self.sector_size]

    def _chain(self, start):
        chain = []
        s = start
        while s != 0xFFFFFFFE and s != 0xFFFFFFFF and s != 0xFFFFFFFC:
            chain.append(s)
            s = self.fat[s]
        return chain

    def _minichain(self, start):
        chain = []
        s = start
        while s != 0xFFFFFFFE:
            chain.append(s)
            s = self.minifat[s]
        return chain

    def read_stream(self, name):
        e = [x for x in self.entries if x['name']==name][0]
        if e['size'] < self.mini_cutoff:
            chain = self._minichain(e['start'])
            data = b''.join(self.ministream[s*self.mini_sector_size:(s+1)*self.mini_sector_size] for s in chain)
        else:
            chain = self._chain(e['start'])
            data = b''.join(self._sector_data(s) for s in chain)
        return data[:e['size']]


with open(IN_PATH, 'rb') as f:
    data = f.read()

# Plain (unencrypted) xlsx files are zip files starting with "PK". Just copy.
if data[:2] == b'PK':
    shutil.copy(IN_PATH, OUT_PATH)
    print('file is not encrypted, copied as-is')
    sys.exit(0)

cfb = CFB(data)

enc_info = cfb.read_stream('EncryptionInfo')
enc_pkg = cfb.read_stream('EncryptedPackage')

xml_data = enc_info[8:]
ns = {'p': 'http://schemas.microsoft.com/office/2006/encryption',
      'pwd': 'http://schemas.microsoft.com/office/2006/keyEncryptor/password'}
root = ET.fromstring(xml_data)
keyData_attrib = root.find('p:keyData', ns).attrib
encryptor = root.find('p:keyEncryptors', ns).find('p:keyEncryptor', ns)
pwdkey = encryptor.find('pwd:encryptedKey', ns)

spinCount = int(pwdkey.attrib['spinCount'])
saltValue = __import__('base64').b64decode(pwdkey.attrib['saltValue'])
hashAlgorithm = pwdkey.attrib['hashAlgorithm']
keyBits = int(pwdkey.attrib['keyBits'])
encryptedVerifierHashInput = __import__('base64').b64decode(pwdkey.attrib['encryptedVerifierHashInput'])
encryptedVerifierHashValue = __import__('base64').b64decode(pwdkey.attrib['encryptedVerifierHashValue'])
encryptedKeyValue = __import__('base64').b64decode(pwdkey.attrib['encryptedKeyValue'])


def h(data, alg=hashAlgorithm):
    if alg.lower() == 'sha512':
        return hashlib.sha512(data).digest()
    elif alg.lower() == 'sha384':
        return hashlib.sha384(data).digest()
    else:
        return hashlib.sha256(data).digest()


iKey = h(saltValue + PASSWORD.encode('utf-16-le'))
for i in range(spinCount):
    iKey = h(struct.pack('<I', i) + iKey)


def derive_key(iKey, blockKeyHex, keyBits):
    blockKey = bytes.fromhex(blockKeyHex)
    k = h(iKey + blockKey)
    keyBytes = keyBits // 8
    if len(k) < keyBytes:
        k = k + b'\x36' * (keyBytes - len(k))
    return k[:keyBytes]


def aes_cbc_decrypt(key, iv, ct):
    cipher = Cipher(algorithms.AES(key), modes.CBC(iv))
    d = cipher.decryptor()
    return d.update(ct) + d.finalize()


verifierInputKey = derive_key(iKey, 'fea7d2763b4b9e79', keyBits)
verifierHashKey = derive_key(iKey, 'd7aa0f6d3061344e', keyBits)
keyValueKey = derive_key(iKey, '146e0be7abacd0d6', keyBits)

hashInput = aes_cbc_decrypt(verifierInputKey, saltValue, encryptedVerifierHashInput)
hashValue = aes_cbc_decrypt(verifierHashKey, saltValue, encryptedVerifierHashValue)
calc_hash = h(hashInput)
hashSize = int(pwdkey.attrib.get('hashSize', len(calc_hash)))
if calc_hash[:hashSize] != hashValue[:hashSize]:
    print('WARNING: password verification failed - decrypted output may be invalid')

secretKey = aes_cbc_decrypt(keyValueKey, saltValue, encryptedKeyValue)
secretKey = secretKey[:keyBits // 8]

pkgSalt = __import__('base64').b64decode(keyData_attrib['saltValue'])
total_size = struct.unpack('<Q', enc_pkg[:8])[0]
body = enc_pkg[8:]
SEG = 4096
out = bytearray()
for i in range(0, len(body), SEG):
    seg = body[i:i+SEG]
    seg_iv = h(pkgSalt + struct.pack('<I', i // SEG))[:16]
    dec = aes_cbc_decrypt(secretKey, seg_iv, seg)
    out.extend(dec)
out = bytes(out[:total_size])
with open(OUT_PATH, 'wb') as f:
    f.write(out)
print('saved', len(out), 'bytes to', OUT_PATH)
