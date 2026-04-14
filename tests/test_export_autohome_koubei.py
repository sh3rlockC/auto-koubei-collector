import importlib.util
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).resolve().parents[1] / "scripts" / "export_autohome_koubei.py"


def load_module():
    spec = importlib.util.spec_from_file_location("export_autohome_koubei", SCRIPT_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class ExportAutohomeKoubeiTests(unittest.TestCase):
    def test_extract_cards_keeps_summary_prefix_with_overall_rating(self):
        module = load_module()
        lines = [
            "    - listitem:",
            "      - link [ref=e142] [nth=9]:",
            "        - /url: https://i.autohome.com.cn/255825840",
            '      - link "Li8260 长安启源A06" [ref=e143]:',
            "        - /url: https://i.autohome.com.cn/255825840",
            "        - paragraph: Li8260",
            "        - paragraph: 长安启源A06",
            "      - text: 关注",
            "      - paragraph: 2026-01-28 发表口碑",
            '      - text: 综合口碑评分 4.57',
            '          - paragraph: "5"',
            "          - paragraph: 空间",
            '          - paragraph: "4"',
            "          - paragraph: 驾驶感受",
            '      - heading "我可以称它为小米su6嘛嘿嘿" [ref=e144] [level=1]:',
            '        - link "我可以称它为小米su6嘛嘿嘿" [ref=e145]:',
            "          - /url: https://k.autohome.com.cn/detail/view_01kbmfsy5n6mvk8e9p6wsg0000.html#pvareaid=2112108",
            '      - link "长安启源A06" [ref=e146] [nth=9]:',
            "        - /url: /8140",
            '      - link "2026款 630Max" [ref=e147]:',
            "        - /url: /spec/75034",
            "      - text: 询底价",
            "        - listitem: 共9张图",
            "        - listitem: 630km 行驶里程",
            "      - text: 满意 关于这款车...",
            "        - listitem:",
            '          - link "44" [ref=e148]:',
            "            - /url: https://k.autohome.com.cn/detail/view_01kbmfsy5n6mvk8e9p6wsg0000.html?pvareaid=6854261#mykoubeiCommentDiv",
            "        - listitem: 收藏",
            "        - listitem:",
            '          - link "查看完整口碑" [ref=e149] [nth=9]:',
            "            - /url: https://k.autohome.com.cn/detail/view_01kbmfsy5n6mvk8e9p6wsg0000.html#pvareaid=2112108",
            "  - text: 相关车系口碑推荐",
        ]

        cards = module.extract_cards(lines)
        self.assertEqual(len(cards), 1)
        row = module.parse_card_summary(cards[0], 2)
        self.assertEqual(row["综合口碑"], "4.57")
        self.assertEqual(row["发表日期"], "2026-01-28")

    def test_normalize_detail_url_removes_query_and_fragment(self):
        module = load_module()
        self.assertEqual(
            module.normalize_detail_url(
                "https://k.autohome.com.cn/detail/view_01abc.html?pvareaid=6854261#mykoubeiCommentDiv"
            ),
            "https://k.autohome.com.cn/detail/view_01abc.html",
        )

    def test_compose_review_text_keeps_full_sections_and_append_review(self):
        module = load_module()
        detail_text = module.compose_review_text(
            [
                {"heading": "最满意", "body": "空间大，底盘稳。"},
                {"heading": "最不满意", "body": "车机还有优化空间。"},
                {"heading": "空间", "body": "后排够大，掀背实用。"},
                {"heading": "", "body": ""},
            ],
            [
                {"label": "购车4个月后追加口碑 | 2026-04-11", "body": "跑了五千多公里，续航焦虑明显降低。"},
            ],
        )
        self.assertIn("最满意：空间大，底盘稳。", detail_text)
        self.assertIn("最不满意：车机还有优化空间。", detail_text)
        self.assertIn("空间：后排够大，掀背实用。", detail_text)
        self.assertIn("购车4个月后追加口碑 | 2026-04-11：跑了五千多公里，续航焦虑明显降低。", detail_text)

    def test_row_from_detail_payload_builds_single_review_row(self):
        module = load_module()
        payload = {
            "title": "国产新能源六边形战士",
            "username": "PLUTO3125",
            "published_at": "2025-12-25",
            "overall_rating": "5",
            "model": "2026款 630激光Ultra",
            "meta_items": [
                "605km 行驶里程",
                "15.5kWh 冬季电耗",
                "14.19万 裸车购买价",
                "2025-12 购买时间",
                "邵阳 购买地点",
            ],
            "sections": [
                {"heading": "最满意", "body": "空间利用率高，底盘扎实。"},
                {"heading": "最不满意", "body": "冬季续航打折。"},
                {"heading": "空间", "body": "后排和后备厢都够用。"},
            ],
            "append_reviews": [
                {"label": "购车4个月后追加口碑 | 2026-04-11", "body": "开了超过5000公里，快充速度满意。"},
            ],
            "source_link": "https://k.autohome.com.cn/detail/view_01kd9ehsme6mvkcdhk6rt00000.html",
        }

        row = module.row_from_detail_payload(payload, page=3)

        self.assertEqual(row["口碑标题"], "国产新能源六边形战士")
        self.assertEqual(row["用户名"], "PLUTO3125")
        self.assertEqual(row["最满意"], "空间利用率高，底盘扎实。")
        self.assertEqual(row["最不满意"], "冬季续航打折。")
        self.assertEqual(row["抓取页码"], 3)
        self.assertEqual(row["来源链接"], payload["source_link"])
        self.assertEqual(row["数据类型"], "车主购车口碑")
        self.assertIn("空间：后排和后备厢都够用。", row["评价详情"])
        self.assertIn("购车4个月后追加口碑 | 2026-04-11：开了超过5000公里，快充速度满意。", row["评价详情"])

    def test_row_from_detail_payload_computes_overall_rating_from_section_scores(self):
        module = load_module()
        payload = {
            "title": "我可以称它为小米su6嘛嘿嘿",
            "username": "Li8260",
            "published_at": "2025-12-04",
            "overall_rating": "",
            "model": "长安启源A06 2026款 630Max",
            "meta_items": [],
            "sections": [
                {"heading": "最满意", "body": "满意内容", "score": ""},
                {"heading": "最不满意", "body": "不满意内容", "score": ""},
                {"heading": "空间", "body": "空间不错", "score": "5"},
                {"heading": "驾驶感受", "body": "驾驶感受不错", "score": "4"},
                {"heading": "续航", "body": "续航不错", "score": "4"},
                {"heading": "外观", "body": "外观不错", "score": "5"},
                {"heading": "内饰", "body": "内饰不错", "score": "5"},
                {"heading": "性价比", "body": "性价比不错", "score": "5"},
                {"heading": "智能化", "body": "智能化不错", "score": "4"},
            ],
            "append_reviews": [],
            "source_link": "https://k.autohome.com.cn/detail/view_01kbmfsy5n6mvk8e9p6wsg0000.html",
        }

        row = module.row_from_detail_payload(payload, page=2)
        self.assertEqual(row["综合口碑"], "4.57")

    def test_row_from_detail_payload_prefers_detail_section_average_over_input_rating(self):
        module = load_module()
        payload = {
            "title": "测试口碑",
            "username": "tester",
            "published_at": "2026-04-13",
            "overall_rating": "5",
            "model": "2026款 630Max",
            "meta_items": [],
            "sections": [
                {"heading": "空间", "body": "空间不错", "score": "5"},
                {"heading": "驾驶感受", "body": "驾驶感受不错", "score": "4"},
                {"heading": "续航", "body": "续航不错", "score": "4"},
                {"heading": "外观", "body": "外观不错", "score": "5"},
                {"heading": "内饰", "body": "内饰不错", "score": "5"},
                {"heading": "性价比", "body": "性价比不错", "score": "5"},
                {"heading": "智能化", "body": "智能化不错", "score": "4"},
            ],
            "append_reviews": [],
            "source_link": "https://k.autohome.com.cn/detail/view_01test.html",
        }

        row = module.row_from_detail_payload(payload, page=1)
        self.assertEqual(row["综合口碑"], "4.57")


if __name__ == "__main__":
    unittest.main()
