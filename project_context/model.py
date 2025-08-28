# Raspberry Pi 4 object detection with better accuracy
# - Model: EfficientDet-Lite2 (TFLite with metadata)
# - Labels: read from TFLite metadata (no hardcoded list)
# - Preprocess: letterbox to preserve aspect ratio
# - Runtime: tflite_runtime if available; fallback to TensorFlow Lite
# - Camera: Picame  ra2; fallback to OpenCV if Picamera2 unavailable

import os, io, time, threading, zipfile, urllib.request
import numpy as np

# Prefer lightweight runtime on Raspberry Pi
try:
    from tflite_runtime.interpreter import Interpreter, load_delegate
except Exception:
    from tensorflow.lite import Interpreter  # fallback
    def load_delegate(*args, **kwargs):  # no-op on TF Lite
        raise OSError("EdgeTPU delegate not available")

# Camera
USE_PICAMERA2 = True
try:
    from picamera2 import Picamera2
except Exception:
    USE_PICAMERA2 = False
    import cv2

import matplotlib.pyplot as plt
import matplotlib.patches as patches

# ---------------------------------------------------------------------
# Model config
# ---------------------------------------------------------------------
MODEL_CHOICE = "efficientdet_lite2"
MODEL_URLS = {
    # Metadata build includes label file
    "efficientdet_lite2": "https://tfhub.dev/tensorflow/lite-model/efficientdet/lite2/detection/metadata/1?lite-format=tflite",
}
MODEL_FILENAME = f"{MODEL_CHOICE}.tflite"

# ---------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------
def download(url, path):
    if os.path.exists(path) and os.path.getsize(path) > 1_000_000:
        return
    print(f"[INFO] downloading {path}")
    urllib.request.urlretrieve(url, path)
    if not os.path.exists(path) or os.path.getsize(path) < 1_000_000:
        raise RuntimeError("download failed or file too small")

def load_labels_from_tflite(tflite_path):
    # TFLite with metadata is a zip; try to read any associated labels file
    labels = None
    try:
        with zipfile.ZipFile(tflite_path, "r") as zf:
            # common names in metadata builds
            cand = [n for n in zf.namelist() if n.lower().endswith((".txt", ".labels", "labelmap"))]
            # pick shortest path that contains 'label'
            cand = sorted([n for n in cand if "label" in n.lower()], key=len)
            if cand:
                with zf.open(cand[0]) as f:
                    txt = f.read().decode("utf-8", errors="ignore")
                labels = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    except zipfile.BadZipFile:
        pass
    return labels or []

def letterbox(img, new_w, new_h):
    h, w = img.shape[:2]
    scale = min(new_w / w, new_h / h)
    nw, nh = int(w * scale), int(h * scale)
    resized = None
    try:
        import cv2
        resized = cv2.resize(img, (nw, nh), interpolation=cv2.INTER_LINEAR)
    except Exception:
        # fallback without OpenCV
        from PIL import Image
        resized = np.array(Image.fromarray(img).resize((nw, nh)))
    canvas = np.zeros((new_h, new_w, 3), dtype=img.dtype)
    top = (new_h - nh) // 2
    left = (new_w - nw) // 2
    canvas[top:top+nh, left:left+nw] = resized
    return canvas, scale, left, top

def quantize_input(x, in_dtype, in_scale, in_zero):
    if in_dtype == np.float32:
        return (x.astype(np.float32) / 255.0).astype(np.float32)
    if in_dtype == np.uint8:
        if in_scale > 0:
            q = (x.astype(np.float32) / 255.0)
            q = (q / in_scale + in_zero).round().clip(0, 255).astype(np.uint8)
            return q
        return x.astype(np.uint8)
    if in_dtype == np.int8:
        if in_scale > 0:
            q = (x.astype(np.float32) / 255.0)
            q = (q / in_scale + in_zero).round().clip(-128, 127).astype(np.int8)
            return q
        # fallback
        q = ((x.astype(np.float32) - 127.5) / 127.5).clip(-1, 1)
        return (q / 0.0078125).round().clip(-128, 127).astype(np.int8)
    return x.astype(in_dtype)

# ---------------------------------------------------------------------
# Setup model
# ---------------------------------------------------------------------
download(MODEL_URLS[MODEL_CHOICE], MODEL_FILENAME)
labels = load_labels_from_tflite(MODEL_FILENAME)

interpreter = Interpreter(model_path=MODEL_FILENAME)
interpreter.allocate_tensors()
input_details  = interpreter.get_input_details()
output_details = interpreter.get_output_details()
_, in_h, in_w, _ = input_details[0]["shape"]
in_dtype = input_details[0]["dtype"]
in_scale, in_zero = (input_details[0].get("quantization", (0.0, 0)) or (0.0, 0))

# ---------------------------------------------------------------------
# Camera init
# ---------------------------------------------------------------------
if USE_PICAMERA2:
    picam2 = Picamera2()
    config = picam2.create_preview_configuration(main={"size": (640, 480), "format": "RGB888"})
    picam2.configure(config)
    picam2.start()
else:
    cap = cv2.VideoCapture(0)
    cap.set(3, 640); cap.set(4, 480)
    if not cap.isOpened():
        raise SystemExit("[ERROR] webcam open failed")

# ---------------------------------------------------------------------
# Inference loop (threaded)
# ---------------------------------------------------------------------
score_thresh = 0.5
frame_lock = threading.Lock()
latest_frame = None
detections = []
fps = 0.0
running = True

def detection_loop():
    global detections, fps, running
    t_prev = time.time()
    while running:
        with frame_lock:
            frame = None if latest_frame is None else latest_frame.copy()
        if frame is None:
            time.sleep(0.005); continue

        # letterbox
        lb_img, scale, left, top = letterbox(frame, in_w, in_h)
        x = quantize_input(lb_img, in_dtype, in_scale or 0.0, in_zero or 0)
        x = np.expand_dims(x, 0)

        interpreter.set_tensor(input_details[0]["index"], x)
        interpreter.invoke()

        boxes   = interpreter.get_tensor(output_details[0]["index"])[0]  # [N,4] ymin,xmin,ymax,xmax in [0,1]
        classes = interpreter.get_tensor(output_details[1]["index"])[0]
        scores  = interpreter.get_tensor(output_details[2]["index"])[0]
        count   = int(interpreter.get_tensor(output_details[3]["index"])[0])

        h, w = frame.shape[:2]
        dets = []
        for i in range(count):
            s = float(scores[i])
            if s < score_thresh: continue
            c = int(classes[i])
            # Map label: prefer 0-based; if that fails, try 1-based
            if 0 <= c < len(labels):
                name = labels[c]
            elif 1 <= c <= len(labels):
                name = labels[c-1]
            else:
                name = f"id:{c}"

            ymin, xmin, ymax, xmax = [float(v) for v in boxes[i]]

            # de-letterbox: model coords -> input pixels -> original frame
            x1i = xmin * in_w; y1i = ymin * in_h
            x2i = xmax * in_w; y2i = ymax * in_h
            x1 = int((x1i - left) / scale); y1 = int((y1i - top) / scale)
            x2 = int((x2i - left) / scale); y2 = int((y2i - top) / scale)

            x1 = max(0, min(w-1, x1)); y1 = max(0, min(h-1, y1))
            x2 = max(0, min(w-1, x2)); y2 = max(0, min(h-1, y2))
            if x2 <= x1 or y2 <= y1: continue

            dets.append(((x1, y1, x2, y2), name, s))
        detections = dets

        t_now = time.time()
        fps = 1.0 / (t_now - t_prev) if t_now > t_prev else 0.0
        t_prev = t_now

# ---------------------------------------------------------------------
# Display
# ---------------------------------------------------------------------
import threading
plt.ion()
fig, ax = plt.subplots(figsize=(8, 6))
thread = threading.Thread(target=detection_loop, daemon=True); thread.start()
print("Running. Close window or Ctrl+C to quit.")

try:
    while running:
        if USE_PICAMERA2:
            frame = picam2.capture_array()
        else:
            ret, frame_bgr = cap.read()
            if not ret: time.sleep(0.01); continue
            frame = frame_bgr[:, :, ::-1]  # BGR->RGB

        with frame_lock:
            latest_frame = frame

        ax.clear(); ax.imshow(frame); ax.set_title(f"{MODEL_CHOICE} - FPS: {fps:.1f}"); ax.axis("off")
        for (x1, y1, x2, y2), name, s in detections:
            rect = patches.Rectangle((x1, y1), x2-x1, y2-y1, linewidth=2, edgecolor="lime", facecolor="none")
            ax.add_patch(rect)
            ax.text(x1, max(0, y1-5), f"{name} {s:.2f}", color="yellow", fontsize=8,
                    bbox=dict(facecolor="black", alpha=0.5, pad=1))
        plt.pause(0.001)
except KeyboardInterrupt:
    pass
finally:
    running = False
    thread.join(timeout=1.0)
    if USE_PICAMERA2:
        try: picam2.stop()
        except Exception: pass
    else:
        cap.release()
    plt.close(fig)
