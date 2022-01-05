import qrcode
import qrcode.image.svg

factory = qrcode.image.svg.SvgPathFillImage

data =  "0\t"\
        "Bysco Technology (Shenzhen) Co., Ltd\t" \
        "BY-SG A020210104\t" \
        "04/01/2021\t" \
        "Sun JinLong\t" \
        "MiniSM50-LB-F1616\t" \
        "100\t" \
        "Wu Shao Hua Ben\t" \
        "05/01/2022\n"

img = qrcode.make(data, image_factory=factory)
# Save svg file somewhere
img.save("qrcode.svg")