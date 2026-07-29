# -*- coding: utf-8 -*-
"""PaddleOCR 한국어 문서 배치 실행기.

금융 문서 이미지를 외부 API로 보내지 않고 로컬 CPU에서만 읽는다. 모델은 작업 저장소가
아닌 사용자 모델 캐시(~/.paddlex)에 설치되며, 한 번 모델을 올린 뒤 여러 이미지를 연속
처리해 이미지마다 40~60초씩 초기화되는 비용을 없앤다.
"""
import argparse
import json
import os
import sys
import traceback

# Windows Paddle 3.3의 oneDNN PIR 변환 오류를 피한다. 정확도 모델 자체는 그대로 쓴다.
os.environ.setdefault("FLAGS_use_mkldnn", "0")
os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="UTF-8 JSON: 이미지 절대경로 배열")
    ap.add_argument("--output", required=True, help="UTF-8 JSON 결과 파일")
    ap.add_argument("--min-score", type=float, default=0.25)
    args = ap.parse_args()

    paths = json.load(open(args.input, encoding="utf-8"))
    if not isinstance(paths, list):
        raise SystemExit("input must be a JSON array")

    from paddleocr import PaddleOCR

    engine = PaddleOCR(
        lang="korean",
        text_detection_model_name="PP-OCRv5_mobile_det",
        text_recognition_model_name="korean_PP-OCRv5_mobile_rec",
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        enable_mkldnn=False,
        device="cpu",
        text_det_limit_side_len=1920,
        text_rec_score_thresh=args.min_score,
    )
    out = {}
    for n, path in enumerate(paths, 1):
        item = {"text": "", "error": "", "lines": 0}
        try:
            lines = []
            for result in engine.predict(path):
                data = result.json
                data = data.get("res", data) if isinstance(data, dict) else {}
                texts = data.get("rec_texts") or []
                scores = data.get("rec_scores") or []
                for i, text in enumerate(texts):
                    score = float(scores[i]) if i < len(scores) else 1.0
                    text = str(text or "").strip()
                    if text and score >= args.min_score:
                        lines.append(text)
            item["text"] = "\n".join(lines)
            item["lines"] = len(lines)
        except Exception as exc:
            item["error"] = f"{type(exc).__name__}: {exc}"[:500]
            traceback.print_exc(file=sys.stderr)
        out[str(path)] = item
        print(f"[{n}/{len(paths)}] {os.path.basename(path)} · {item['lines']}줄", flush=True)

    tmp = args.output + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp, args.output)


if __name__ == "__main__":
    main()
