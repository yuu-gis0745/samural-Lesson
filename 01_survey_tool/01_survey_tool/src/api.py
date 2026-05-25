from pathlib import Path
from fastapi import FastAPI
from importlib.machinery import SourceFileLoader

# プロジェクトの基準フォルダ
base_dir = Path(r"E:\protopype\01_survey_tool")

# 03_check_forest_survey2.py を読み込む
checker = SourceFileLoader(
    "checker",
    str(base_dir / "src" / "03_check_forest_survey2.py")
).load_module()

app = FastAPI()


@app.get("/")
def read_root():
    return {"message": "Survey check API is running"}


@app.post("/check")
def check_survey():
    result = checker.run_check(
        input_file_path=str(base_dir / "01_input" / "sample_forest_survey_3plots.xlsx"),
        output_dir_path=str(base_dir / "03_output"),
        cell_mapping_file_path=str(base_dir / "02_master" / "forest_survey_cell_mapping.xlsx"),
        check_rules_file_path=str(base_dir / "02_master" / "check_rules_forest_survey.csv"),)

    return result