# -*- coding: utf-8 -*-
"""
达成度报告生成器 - 配置参数
"""

from dataclasses import dataclass


@dataclass
class Config:
    """配置参数类"""
    # 达成度目标占比（%）
    ratio_1: int = 50  # 目标一占比
    ratio_2: int = 30  # 目标二占比
    ratio_3: int = 20  # 目标三占比

    # 成绩占比（%）
    regular_score_ratio: int = 30  # 平时成绩占比
    final_score_ratio: int = 70    # 期末成绩占比

    # 达成度期望值
    achievement_expectation: float = 0.6

    def validate(self) -> tuple[bool, str]:
        """验证配置参数

        Returns:
            (是否有效, 错误信息)
        """
        # 验证目标占比之和为100
        if self.ratio_1 + self.ratio_2 + self.ratio_3 != 100:
            return False, f"目标占比之和必须为100%，当前为{self.ratio_1 + self.ratio_2 + self.ratio_3}%"

        # 验证成绩占比之和为100
        if self.regular_score_ratio + self.final_score_ratio != 100:
            return False, f"成绩占比之和必须为100%，当前为{self.regular_score_ratio + self.final_score_ratio}%"

        # 验证各项占比为非负数
        if any(x < 0 for x in [self.ratio_1, self.ratio_2, self.ratio_3,
                               self.regular_score_ratio, self.final_score_ratio]):
            return False, "所有占比必须为非负数"

        # 验证达成度期望值在0-1之间
        if not 0 <= self.achievement_expectation <= 1:
            return False, f"达成度期望值必须在0-1之间，当前为{self.achievement_expectation}"

        return True, ""
