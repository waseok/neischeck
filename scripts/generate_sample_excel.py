from pathlib import Path

import pandas as pd


def main() -> None:
    df = pd.DataFrame(
        [
            {"학번": "10101", "이름": "홍길동", "세특": "유튜브 자료를 참고하여 발표함."},
            {"학번": "10102", "이름": "김학생", "세특": "교내 수업에서 탐구활동을 성실히 수행함."},
            {"학번": "10103", "이름": "박학생", "세특": "모의고사 성적 향상을 위해 노력함."},
        ]
    )
    out = Path("data") / "sample_input.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(out, index=False)
    print(f"샘플 파일 생성: {out}")


if __name__ == "__main__":
    main()
