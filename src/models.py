"""
Pydantic models for AI review results
"""
from pydantic import BaseModel, Field


class ReviewRow(BaseModel):
    """AI評審結果の1行を表すモデル"""
    requirement_no: str = Field(
        description="要求No（例: REQ-001）"
    )
    requirement_content: str = Field(
        description="要求内容（ペルソナ: 指令）の詳細"
    )
    evaluation: str = Field(
        description="評価結果（〇/△/×のいずれか）"
    )
    compliance_location: str = Field(
        description="適合または不適合の箇所の参照"
    )
    compliance_reason: str = Field(
        description="適合または不適合と判断した理由"
    )
    correction_plan: str = Field(
        description="修正案の内容（ゴールデンケースを含む）"
    )
    response_status: str = Field(
        description="対応の有無"
    )
    response_method: str = Field(
        description="対応方法または非対応の理由"
    )


class ReviewTable(BaseModel):
    """AI評審結果の表全体"""
    rows: list[ReviewRow] = Field(
        description=(
            "評審結果のすべての行。**絶対に1行も省略してはいけない。**"
            "PDFに含まれるすべての要求項目を完全に抽出し、出力すること。"
            "途中で切らずに、最後の行まで必ず含めること。"
        )
    )
