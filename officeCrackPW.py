import msoffcrypto
import io
import itertools
import string
import multiprocessing


# ==================== 密码生成器函数 ====================
def generate_passwords(user_config,common_passwords ,bases_words = None,leet_list =None):
    """
    一个生成器函数，按策略生成密码
    可以按需扩展更多规则
    """


    # === 1. 常见弱密码（高频尝试）
    common_passwords = common_passwords or [
            "123456", "password", "123456789", "12345678", "12345", "1234567",
            "admin", "password123", "excel", "123", "000000", "test", "pass",
            "abc123", "111111", "qwerty"
        ]
    for pwd in common_passwords:
        yield pwd

    # === 2. 基础词 + 年份/数字后缀（掩码攻击）
    bases = bases_words or ["admin", "user", "password", "excel", "sheet", "doc"]
    years = [str(year) for year in range(2000, 2026)]
    numbers = [str(i).zfill(n) for n in [1,2,3] for i in range(0, 10**n)]

    for base in bases:
        for suffix in years + numbers:
            yield base + suffix
            yield base.capitalize() + suffix

    # === 3. Leet Speak 变体（e.g. P@ssw0rd）
    def leet_variants(word):
        leet_map = {
            'a': ['a', 'A', '@', '4'],
            'b': ['b', 'B', '8'],
            'c': ['c', 'C', '('],
            'd': ['d','D'],
            'e': ['e', 'E', '3'],
            'f': ['f',"F"],
            'g': ['g', 'G', '6'],
            'h': ['h','H'],
            'i': ['i', 'I', '1', '!'],
            'j': ['j','J'],
            'k': ['k','K'],
            'l': ['l', 'L', '1'],
            'm': ['m','M','nn'],
            'n': ['n','N'],
            'o': ['o', 'O', '0'],
            'p': ['P', 'p'],
            'q': ['q', 'Q','9'],
            'r': ['R','r'],
            's': ['s', 'S', '$', '5'],
            't': ['t', 'T', '7'],
            'u': ['U','u'],
            'v': ['v','V',"\/"],
            'w': ['W',"w",'vv'],
            'x': ['x','X'],
            'y': ['Y','y'],
            'z': ['z', 'Z', '2'],
        }
        chars = [leet_map.get(c.lower(), [c]) for c in word]
        for combo in itertools.product(*chars):
            yield ''.join(combo)

    leet_list = leet_list or ["pass", "password", "excel"]
    for base in leet_list:
        for variant in leet_variants(base):
            yield variant
            yield variant + "123"
            yield variant + "2025"

    # === 4. 简单暴力破解（长度 4-6，小写字母+数字）
    charset = string.ascii_lowercase + string.digits
    for length in range(int(user_config["BruteForceLengthStart"]), int(user_config["BruteForceLengthEnd"])):
        for pwd_tuple in itertools.product(charset, repeat=length):
            yield ''.join(pwd_tuple)


# ==================== 使用密码生成器破解 Excel ====================

def crack_excel_with_generator(filename,user_config,common_passwords = None,bases_words = None,leet_list =None):
    """
    使用生成器生成密码，尝试破解加密的 Excel 文件
    """
    with open(filename, "rb") as f:
        office_file = msoffcrypto.OfficeFile(f)

        print(f"开始破解 {filename} ...")

        for i, password in enumerate(generate_passwords(user_config,common_passwords,bases_words,leet_list)):
            office_file.load_key(password=password)
            decrypted = io.BytesIO()
            try:
                office_file.decrypt(decrypted)  # 成功则无异常
                print(f"[+] 破解成功！密码: {password}")

                try:
                    # 可选：保存解密文件
                    with open("decrypted_output.xlsx", "wb") as out:
                        out.write(decrypted.getvalue())
                    return password
                except Exception as e:
                    print("! 無法保存解密文件，decrypted_output.xlsx可能己被佔用")
                    return password

            except Exception as e:
                pass

            # 进度提示
            if (i + 1) % 100 == 0:
                print(f"[-] 已尝试 {i + 1} 个密码，最近: {password}")


        print("[-] 在生成的密码中未找到正确密码。")
    return None


# ==================== 运行示例 ====================

#if __name__ == "__main__":
#    crack_excel_with_generator("test.xlsx")
