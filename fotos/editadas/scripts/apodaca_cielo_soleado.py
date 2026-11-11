import cv2, numpy as np, sys
src, skysrc, out, dbg = sys.argv[1:5]
img = cv2.imread(src); h, w = img.shape[:2]
hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
cand = ((hsv[...,1] < 45) & (hsv[...,2] > 150)).astype(np.uint8)
cand[int(h*0.62):] = 0
n, lab = cv2.connectedComponents(cand)
top_labels = set(np.unique(lab[0][cand[0] > 0])) - {0}
sky = np.isin(lab, list(top_labels)).astype(np.uint8) * 255
sky = cv2.morphologyEx(sky, cv2.MORPH_CLOSE, np.ones((7,7), np.uint8))
sky = cv2.morphologyEx(sky, cv2.MORPH_OPEN, np.ones((3,3), np.uint8))
cv2.imwrite(dbg, sky)
# cielo nuevo (de la foto de Escobedo, misma ciudad)
s = cv2.imread(skysrc)[0:300, 900:1600]
s = cv2.resize(s, (w, int(300 * w / 700)), interpolation=cv2.INTER_CUBIC)
newsky = np.zeros_like(img); hh = min(h, s.shape[0]); newsky[:hh] = s[:hh]
if hh < h: newsky[hh:] = s[-1]
alpha = cv2.GaussianBlur(sky.astype(np.float32)/255, (5,5), 0)[..., None]
res = img.astype(np.float32) * (1 - alpha) + newsky.astype(np.float32) * alpha
# luz de día: más brillo, contraste y calidez fuera del cielo
ground = 1 - alpha
lab_ = cv2.cvtColor(np.clip(res,0,255).astype(np.uint8), cv2.COLOR_BGR2LAB).astype(np.float32)
L = lab_[...,0]
L2 = np.clip((L - 128) * 1.12 + 128 + 10, 0, 255)
lab_[...,0] = L * alpha[...,0] + L2 * ground[...,0]
lab_[...,2] = lab_[...,2] + 5 * ground[...,0]   # cálido (b+)
lab_[...,1] = lab_[...,1] + 1 * ground[...,0]
res = cv2.cvtColor(np.clip(lab_,0,255).astype(np.uint8), cv2.COLOR_LAB2BGR).astype(np.float32)
# charcos: zonas claras y grises del piso toman un leve tono del cielo
hsv2 = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
puddle = ((hsv2[...,1] < 40) & (hsv2[...,2] > 165)).astype(np.float32)
puddle[:int(h*0.775)] = 0
puddle = cv2.GaussianBlur(puddle, (9,9), 0)[..., None] * 0.5
tint = np.array([200, 160, 110], np.float32)  # azul cielo en BGR
res = res * (1 - puddle) + tint * puddle
# cielo bajo el pórtico (cerca del horizonte): azul pálido, sólo píxeles claros y grises de ese recuadro
box = np.zeros((h,w), np.float32)
sub = ((hsv[330:425, 1005:1218, 1] < 35) & (hsv[330:425, 1005:1218, 2] > 175)).astype(np.float32)
box[330:425, 1005:1218] = sub
box = cv2.GaussianBlur(box, (5,5), 0)[..., None] * 0.55
pale = np.array([235, 205, 170], np.float32)
res = res * (1 - box) + pale * box
cv2.imwrite(out, np.clip(res,0,255).astype(np.uint8), [cv2.IMWRITE_JPEG_QUALITY, 93])
