# fix_cert.py
# 这是一个生成 "完美" 自签名证书的脚本
import os
import datetime
import ipaddress
from cryptography import x509
from cryptography.x509.oid import NameOID
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization

# 确保证书目录存在
if not os.path.exists('ssl'):
    os.makedirs('ssl')

def generate_robust_cert():
    # 1. 生成私钥
    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    # 2. 设置证书的基本信息 (CN)
    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, u"127.0.0.1"),
    ])

    # 3. 关键步骤：添加 SAN (主题备用名称)
    # 现代浏览器只认这个字段，必须同时包含 localhost 和 127.0.0.1
    alt_names = [
        x509.DNSName(u"localhost"),
        x509.IPAddress(ipaddress.IPv4Address(u"127.0.0.1"))
    ]
    
    # 4. 构建证书
    cert = x509.CertificateBuilder().subject_name(
        subject
    ).issuer_name(
        issuer
    ).public_key(
        key.public_key()
    ).serial_number(
        x509.random_serial_number()
    ).not_valid_before(
        datetime.datetime.utcnow()
    ).not_valid_after(
        # 有效期 10 年
        datetime.datetime.utcnow() + datetime.timedelta(days=3650)
    ).add_extension(
        x509.SubjectAlternativeName(alt_names),
        critical=False,
    ).sign(key, hashes.SHA256())

    # 5. 写入文件 (覆盖旧的)
    with open("ssl/server.crt", "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))
    
    with open("ssl/server.key", "wb") as f:
        f.write(key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.TraditionalOpenSSL,
            encryption_algorithm=serialization.NoEncryption(),
        ))
        
    print("✅ 全新证书已生成！位置: ssl/server.crt")

if __name__ == "__main__":
    try:
        generate_robust_cert()
    except ImportError:
        print("❌ 缺少依赖库，请先运行: pip install cryptography")