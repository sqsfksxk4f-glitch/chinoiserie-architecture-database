"""
Test Baidu image search and DS verify workflow image steps.
Run: python3 tests/test_baidu_image.py
"""
import sys
import os
import json
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import database as db
from ui.organize_tab import VerificationPanel


class MockFrame:
    """Mock Tk frame for VerificationPanel init."""
    def __init__(self):
        self.children = {}
    def winfo_children(self):
        return []
    def after(self, ms, cb=None):
        return None
    def update_idletasks(self):
        pass
    def wait_window(self, w):
        pass


class MockTree:
    def get_children(self):
        return []
    def item(self, *args):
        return {}
    def selection_set(self, *args):
        pass
    def see(self, *args):
        pass


class MockVar:
    def __init__(self, value=""):
        self._value = value
    def get(self):
        return self._value
    def set(self, v):
        self._value = v


class MockCase:
    def __init__(self):
        self.id = None
        self.name = ""
        self.country = ""
        self.city = ""
        self.image_path = ""

    def get(self, key, default=None):
        return getattr(self, key, default)


def create_mock_vp():
    """Create a minimal VerificationPanel for testing image methods."""
    vp = VerificationPanel.__new__(VerificationPanel)
    vp.frame = MockFrame()
    vp.queue_tree = MockTree()
    vp.overall_var = MockVar()
    vp.paste_text = type('MockText', (), {'delete': lambda *a: None, 'configure': lambda *a, **kw: None})()
    vp.name_result_text = type('MockText', (), {'delete': lambda *a: None, 'configure': lambda *a, **kw: None})()
    vp.other_result_text = type('MockText', (), {'delete': lambda *a: None, 'configure': lambda *a, **kw: None})()
    vp._current_case = None
    vp._image_preview_frame = MockFrame()
    vp._undo_btn = type('Mock', (), {'config': lambda *a, **kw: None})()
    vp._download_btn = type('Mock', (), {'config': lambda *a, **kw: None})()
    vp._active_field = None
    vp._active_mode = "verify"
    vp._verifying_field = None
    vp._verify_field_label = type('Mock', (), {'config': lambda *a, **kw: None})()
    vp._active_field_label = type('Mock', (), {'config': lambda *a, **kw: None})()
    vp._field_statuses = {}
    vp._field_rows = {}
    return vp


class TestBaiduImageFetch(unittest.TestCase):
    """Test the Baidu image fetch API directly."""

    def test_fetch_returns_urls(self):
        """Test that _fetch_baidu_image_urls returns valid (thumb, dl_url) tuples."""
        vp = create_mock_vp()
        results = vp._fetch_baidu_image_urls("佛光寺 建筑", count=10)
        self.assertGreater(len(results), 0, "Should return at least 1 image")
        for thumb, dl_url in results:
            self.assertTrue(thumb.startswith("https://"), f"thumb should be https URL: {thumb[:60]}")
            self.assertTrue(dl_url.startswith("https://"), f"dl_url should be https URL: {dl_url[:60]}")
        print(f"  [OK] Fetched {len(results)} images")
        for i, (thumb, dl) in enumerate(results[:5]):
            print(f"  [{i+1}] thumb: {thumb[:70]}...")
            print(f"       dl:   {dl[:80]}...")

    def test_fetch_with_empty_query(self):
        """Test that empty query returns empty list gracefully."""
        vp = create_mock_vp()
        results = vp._fetch_baidu_image_urls("", count=5)
        self.assertEqual(len(results), 0, "Empty query should return empty list")
        print("  [OK] Empty query returns empty list")

    def test_fetch_with_english_query(self):
        """Test with English query for overseas Chinese buildings."""
        vp = create_mock_vp()
        results = vp._fetch_baidu_image_urls("Chinese temple Bangkok", count=5)
        # May or may not return results depending on Baidu index
        print(f"  [INFO] English query returned {len(results)} images")
        for thumb, dl_url in results[:3]:
            self.assertTrue(thumb.startswith("https://"))
            self.assertTrue(dl_url.startswith("https://"))


class TestImageDownload(unittest.TestCase):
    """Test image download and DB storage."""

    def setUp(self):
        db.init_db()
        # Create a test raw case
        self.test_case_id = db.insert_raw_case({
            "name": "测试百度搜图",
            "country": "日本",
            "city": "东京",
        })
        self.vp = create_mock_vp()
        self.vp._current_case = db.get_raw_case_by_id(self.test_case_id)
        self.vp._current_case["image_path"] = ""
        # Override _refresh_image_preview to avoid Tk widget creation
        self._refresh_called = False
        def mock_refresh(case):
            self._refresh_called = True
        self.vp._refresh_image_preview = mock_refresh

    def test_download_single_image(self):
        """Test downloading a single image from a known URL."""
        vp = self.vp
        test_url = "https://qcloud.dpfile.com/pc/fnXIb7tqgQSRiV5mP8SDdm24WUQCJEpUNxYnC2W8DW_oDaSWvqnh_75XjQjKtRzN.jpg"
        vp._download_selected_images([test_url])

        # Verify image was saved
        case = db.get_raw_case_by_id(self.test_case_id)
        image_path = case.get("image_path", "")
        self.assertTrue(image_path, "image_path should not be empty after download")
        refs = json.loads(image_path)
        self.assertEqual(len(refs), 1, "Should have 1 image ref")
        self.assertTrue(refs[0].startswith("dbimg://"), f"Should be dbimg:// ref: {refs[0]}")

        # Verify image file exists on disk
        img_id = int(refs[0].replace("dbimg://", ""))
        img_rec = db.get_case_image(img_id)
        self.assertIsNotNone(img_rec, "Image record should exist")
        self.assertGreater(len(img_rec.get("image_data", b"")), 0, "Image data should not be empty")
        print(f"  [OK] Downloaded image: {len(img_rec['image_data'])} bytes, file: {img_rec.get('file_path', 'N/A')}")

    def test_download_multiple_images(self):
        """Test downloading multiple images."""
        vp = self.vp
        urls = [
            "https://qcloud.dpfile.com/pc/fnXIb7tqgQSRiV5mP8SDdm24WUQCJEpUNxYnC2W8DW_oDaSWvqnh_75XjQjKtRzN.jpg",
            "https://qcloud.dpfile.com/pc/6ya3J1i9MB6ktSc4iSjoI8d6kF3P_MJpKFBCoDCtIB5MQtWFyszhrrt60YbQhGhj.jpg",
        ]
        vp._download_selected_images(urls)

        case = db.get_raw_case_by_id(self.test_case_id)
        image_path = case.get("image_path", "")
        refs = json.loads(image_path) if image_path else []
        self.assertGreaterEqual(len(refs), 1, "Should download at least 1 image")
        print(f"  [OK] Downloaded {len(refs)} images")

        for ref in refs:
            img_id = int(ref.replace("dbimg://", ""))
            img_rec = db.get_case_image(img_id)
            self.assertIsNotNone(img_rec)
            self.assertGreater(len(img_rec.get("image_data", b"")), 0)

    def test_download_invalid_url(self):
        """Test that invalid URL doesn't crash."""
        vp = self.vp
        vp._download_selected_images(["https://invalid.example.com/not-an-image.jpg"])
        # Should not crash, just download 0
        case = db.get_raw_case_by_id(self.test_case_id)
        print(f"  [OK] Invalid URL handled gracefully, images: {case.get('image_path', 'none')}")


class TestBaiduImageFullFlow(unittest.TestCase):
    """Test the full Baidu image search → download flow without GUI dialog."""

    def setUp(self):
        db.init_db()
        self.test_case_id = db.insert_raw_case({
            "name": "佛光寺",
            "country": "中国",
            "city": "山西省忻州市",
        })
        self.vp = create_mock_vp()
        self.vp._current_case = db.get_raw_case_by_id(self.test_case_id)
        self.vp._current_case["image_path"] = ""
        # Override _refresh_image_preview to avoid Tk widget creation
        def mock_refresh(case):
            pass
        self.vp._refresh_image_preview = mock_refresh

    def test_full_flow_fetch_and_download_first_two(self):
        """Fetch from Baidu API, auto-download first 2 results."""
        vp = self.vp
        results = vp._fetch_baidu_image_urls("佛光寺 中国", count=10)
        self.assertGreater(len(results), 0, "Should fetch images from Baidu")

        # Auto-select first 2 download URLs
        dl_urls = [dl for _, dl in results[:2]]
        vp._download_selected_images(dl_urls)

        case = db.get_raw_case_by_id(self.test_case_id)
        image_path = case.get("image_path", "")
        refs = json.loads(image_path) if image_path else []
        self.assertGreater(len(refs), 0, "Should have downloaded images")
        print(f"  [OK] Full flow: fetched {len(results)}, downloaded {len(refs)} images")

        for ref in refs:
            img_id = int(ref.replace("dbimg://", ""))
            img_rec = db.get_case_image(img_id)
            print(f"  Image {img_id}: {len(img_rec.get('image_data', b''))} bytes")


class TestDsVerifyImageSteps(unittest.TestCase):
    """Test that DsVerifySession steps 5-6 call the right methods."""

    def setUp(self):
        db.init_db()
        self.test_case_id = db.insert_raw_case({
            "name": "测试建筑",
            "country": "日本",
        })

    def test_step_5_calls_baidu_image_search(self):
        """Verify _step_5 uses _baidu_image_search not _select_local_image."""
        import ui.workflows as wf
        # Check the source code of _step_5
        import inspect
        src = inspect.getsource(wf.DsVerifySession._step_5)
        self.assertIn("_baidu_image_search", src, "_step_5 should call _baidu_image_search")
        self.assertNotIn("_select_local_image", src, "_step_5 should NOT call _select_local_image")
        print("  [OK] _step_5 uses _baidu_image_search()")

    def test_step_6_also_uses_baidu_image_search(self):
        """Verify _step_6 uses _baidu_image_search for repeat."""
        import ui.workflows as wf
        import inspect
        src = inspect.getsource(wf.DsVerifySession._step_6)
        self.assertIn("_baidu_image_search", src, "_step_6 should call _baidu_image_search")
        print("  [OK] _step_6 uses _baidu_image_search()")

    def test_baidu_image_search_method_exists(self):
        """Verify _baidu_image_search exists on VerificationPanel."""
        vp = create_mock_vp()
        self.assertTrue(hasattr(vp, "_baidu_image_search"), "VerificationPanel should have _baidu_image_search")
        self.assertTrue(callable(vp._baidu_image_search))
        print("  [OK] _baidu_image_search method exists")

    def test_all_methods_exist(self):
        """Verify all new methods exist."""
        vp = create_mock_vp()
        for method in ["_fetch_baidu_image_urls", "_show_baidu_image_dialog",
                        "_download_selected_images", "_baidu_image_search"]:
            self.assertTrue(hasattr(vp, method), f"Missing method: {method}")
            self.assertTrue(callable(getattr(vp, method)), f"{method} should be callable")
        print("  [OK] All 4 new methods exist on VerificationPanel")

    def test_search_active_field_routes_image_to_baidu(self):
        """Verify _search_active_field routes image_urls+baidu to _baidu_image_search."""
        import inspect
        from ui.organize_tab import VerificationPanel
        src = inspect.getsource(VerificationPanel._search_active_field)
        self.assertIn("field_name == \"image_urls\"", src,
                      "_search_active_field should check for image_urls field")
        self.assertIn("_baidu_image_search", src,
                      "_search_active_field should call _baidu_image_search for image field")
        print("  [OK] _search_active_field routes image_urls to Baidu image search")

    def test_baidu_search_button_exists(self):
        """Verify '百度搜图' button code exists in organize_tab.py."""
        import inspect
        import ui.organize_tab as ot
        src = inspect.getsource(ot)
        self.assertIn("百度搜图", src, "organize_tab.py should contain '百度搜图' button")
        print("  [OK] '百度搜图' button exists in organize_tab")


if __name__ == "__main__":
    unittest.main(verbosity=2)
