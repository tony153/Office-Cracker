import os
import re
import base64
import binascii


def find_data(data: bytes, start_tag: bytes, end_tag: bytes) -> bytes:
    """
    从二进制数据中提取指定标签之间的内容（不含标签本身）
    """
    start = data.find(start_tag)
    end = data.find(end_tag)
    if start == -1 or end == -1:
        raise ValueError(f"无法找到标签: {start_tag} 或 {end_tag}")
    return data[start + len(start_tag):end]


def extract_attribute(data: bytes, attr_name: bytes) -> bytes:
    """
    使用正则提取 XML 属性值，如 saltValue="xxx" → b'xxx'
    """
    pattern = re.escape(attr_name) + b'="([^"]+)"'
    match = re.search(pattern, data)
    if not match:
        raise ValueError(f"无法提取属性: {attr_name}")
    return match.group(1)


def decode_base64_to_hex(b64_str: str) -> str:
    """
    将 Base64 字符串解码为十六进制字符串
    """
    decoded_bytes = base64.b64decode(b64_str)
    return binascii.hexlify(decoded_bytes).decode('ascii').lower()


def extract_office_hash(filename: str) -> str:
    """
    从加密的 Office 文件中提取 John the Ripper 可用的哈希字符串
    支持 Office 2010 (SHA1) 和 Office 2013+ (SHA512)
    """
    # 读取文件前部（足够包含加密元数据）
    with open(filename, 'rb') as f:
        header = f.read(81920)  # 通常足够

    # 查找加密 XML 元数据
    xml_start = b'<?xml version="1.0"'
    encryption_end = b'</encryption>'

    start_pos = header.find(xml_start)
    end_pos = header.find(encryption_end)
    if start_pos == -1 or end_pos == -1:
        raise ValueError("未找到加密元数据（<encryption> XML）")

    xml_metadata = header[start_pos:end_pos + len(encryption_end)]

    # 提取 <keyData> 中的信息
    key_data = find_data(xml_metadata, b"<keyData", b"/>")
    salt_size = int(extract_attribute(key_data, b'saltSize'))
    block_size = int(extract_attribute(key_data, b'blockSize'))  # 虽未用于 John，但可验证
    key_bits = int(extract_attribute(key_data, b'keyBits'))
    hash_algo = extract_attribute(key_data, b'hashAlgorithm').decode('ascii')

    # 确定 Office 版本
    if hash_algo == "SHA1":
        version = 2010
    elif hash_algo == "SHA512":
        version = 2013
    else:
        raise ValueError(f"不支持的哈希算法: {hash_algo}")

    # 提取 <p:encryptedKey> 中的信息
    key_encryptor = find_data(xml_metadata, b"<p:encryptedKey", b"</keyEncryptor>")
    spin_count = int(extract_attribute(key_encryptor, b'spinCount'))
    salt_value_b64 = extract_attribute(key_encryptor, b'saltValue').decode('ascii')
    encrypted_verifier_hash_input_b64 = extract_attribute(key_encryptor, b'encryptedVerifierHashInput').decode('ascii')
    encrypted_verifier_hash_value_b64 = extract_attribute(key_encryptor, b'encryptedVerifierHashValue').decode('ascii')

    # 转换为 hex（John the Ripper 所需格式）
    salt_hex = decode_base64_to_hex(salt_value_b64)
    verifier_hash_input_hex = decode_base64_to_hex(encrypted_verifier_hash_input_b64)
    encrypted_verifier_hash_value_hex = decode_base64_to_hex(encrypted_verifier_hash_value_b64)

    # 构造 John the Ripper 哈希格式
    # 格式: $office$*version*spincount*keybits*saltsize*salt$verifier_input$encrypted_verifier_value
    # 注意：encrypted_verifier_value 取前 64 字符（32 字节）
    john_hash = (
        f"{os.path.basename(filename)}:$office$*{version}*{spin_count}*{key_bits}*"
        f"{salt_size}*{salt_hex}*{verifier_hash_input_hex}*"
        f"{encrypted_verifier_hash_value_hex[:64]}"
    )

    return john_hash


# # ============ 使用示例 ============
# if __name__ == "__main__":
#     try:
#         hash_str = extract_office_hash("ensheet.xlsx")
#         print("John the Ripper 可用哈希:")
#         print(hash_str)
#     except Exception as e:
#         print(f"[错误] {e}")