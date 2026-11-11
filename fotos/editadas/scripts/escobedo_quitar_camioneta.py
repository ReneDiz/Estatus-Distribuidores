import cv2, numpy as np, sys
src, out = sys.argv[1], sys.argv[2]
img = cv2.imread(src).astype(np.float32); h, w = img.shape[:2]
poly = np.array([(1150,1200),(1155,985),(1235,940),(1318,852),(1400,833),(1460,823),(1600,818),(1600,1200)], np.int32)
mask = np.zeros((h,w), np.uint8); cv2.fillPoly(mask,[poly],255)
reg = img[790:870, 1290:1600]
red = ((reg[...,2] - reg[...,1]) > 18) & ((reg[...,2] - reg[...,0]) > 8)
mask[790:870, 1290:1600][red] = 255
mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((9,9),np.uint8))
mask = cv2.dilate(mask, np.ones((7,7),np.uint8))
SH = 470
fill = img.copy()
ys, xs = np.where(mask>0)
sy = np.where(ys >= 940, ys, np.minimum(h-1, ys + 130))
fill[ys, xs] = img[sy, xs - SH]
# 1) corrección por fila (referencia: piso visible a la izquierda)
corr = {}
for y in range(860, h):
    idx = np.where(mask[y] > 0)[0]
    if len(idx) == 0: continue
    x0 = idx[0]
    corr[y] = (img[y, x0-60:x0-5].mean(axis=0) - fill[y, x0:x0+60].mean(axis=0)) * 0.8
for y in range(780, h):
    idx = np.where(mask[y] > 0)[0]
    if len(idx): fill[y, idx] += corr.get(y, corr[860])
# 2) corrección por columna desde el borde superior (empata con piso o sombra de las unidades)
colcorr = np.zeros((w,3),np.float32); top = np.full(w, -1)
for x in range(1150, w):
    c = np.where(mask[:, x] > 0)[0]
    if len(c) == 0: continue
    t = c[0]; top[x] = t
    colcorr[x] = np.median(img[t-10:t-2, x], axis=0) - fill[t+2:t+10, x].mean(axis=0)
colcorr = np.clip(colcorr, -45, 12)
colcorr[1150:] = cv2.GaussianBlur(colcorr[1150:].reshape(-1,1,3), (1,121), 0).reshape(-1,3)
for x in range(1150, w):
    t = top[x]
    if t < 0: continue
    for k in range(0, 40):
        y = t + k
        if y < h and mask[y, x]: fill[y, x] += colcorr[x] * (1 - k/40) ** 2
alpha = cv2.GaussianBlur(mask.astype(np.float32)/255, (15,15), 0)[..., None]
res = np.clip(img*(1-alpha) + fill*alpha, 0, 255).astype(np.uint8)
cv2.imwrite(out, res, [cv2.IMWRITE_JPEG_QUALITY, 92])
