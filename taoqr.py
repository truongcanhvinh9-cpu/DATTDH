import qrcode

# Nội dung QR
data = "bacninh"

# Tạo QR
img = qrcode.make(data)

# Lưu ảnh
img.save("bacninh.png")

print("Đã tạo QR")
