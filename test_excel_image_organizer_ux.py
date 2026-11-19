from __future__ import annotations

import tempfile
import unittest
import zipfile
import gzip
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree

from openpyxl import load_workbook
from PIL import Image, ImageDraw

from excel_image_organizer_ux import (
    DATA_START_ROW,
    GenerationCancelled,
    IMAGE_COLUMN_WIDTH,
    IMAGE_ROW_HEIGHT,
    analyze_selection,
    create_excel,
    excel_column_width_to_pixels,
    excel_row_height_to_pixels,
    extract_group,
    load_settings,
    organize_paths,
    save_settings,
    sequence_sort_key,
    split_group_and_sequence,
    validate_image_cell_size,
    validate_output_filename,
)


class OrganizerTest(unittest.TestCase):
    def test_group_extraction(self) -> None:
        self.assertEqual(extract_group("1-01-1.tif"), "1")
        self.assertEqual(extract_group("No1-3.png"), "No1")
        self.assertEqual(extract_group("2_02_1.jpg"), "2")
        self.assertEqual(extract_group("No12photo3.png"), "No12")
        self.assertEqual(split_group_and_sequence("1-01-2.tif"), ("1", "01-2"))
        self.assertEqual(split_group_and_sequence("No1사진2.png"), ("No1", "사진2"))

    def test_any_non_alphanumeric_character_can_be_group_boundary(self) -> None:
        examples = {
            "No1.01.png": ("No1", "01"),
            "No1#02.png": ("No1", "02"),
            "No1+03.png": ("No1", "03"),
            "No1(04).png": ("No1", "04)"),
            "No1@05.png": ("No1", "05"),
        }
        for filename, expected in examples.items():
            with self.subTest(filename=filename):
                self.assertEqual(split_group_and_sequence(filename), expected)

        paths = [
            Path("No2#01.png"),
            Path("No1(10).png"),
            Path("No1@3.png"),
            Path("No1.02.png"),
            Path("No2+02.png"),
        ]
        groups = organize_paths(paths)
        self.assertEqual([name for name, _ in groups], ["No1", "No2"])
        self.assertEqual(
            [item.path.name for item in groups[0][1]],
            ["No1.02.png", "No1@3.png", "No1(10).png"],
        )
        self.assertEqual(
            [item.path.name for item in groups[1][1]],
            ["No2#01.png", "No2+02.png"],
        )

    def test_repeated_group_boundary_characters_are_removed_safely(self) -> None:
        self.assertEqual(split_group_and_sequence("No1-_+#02.png"), ("No1", "02"))

    def test_windows_output_filename_validation(self) -> None:
        self.assertEqual(validate_output_filename("검사 결과.xlsx"), "검사 결과.xlsx")
        for invalid in ("", "bad:name.xlsx", "folder/name.xlsx", "CON.xlsx", "result. "):
            with self.subTest(filename=invalid):
                with self.assertRaises(ValueError):
                    validate_output_filename(invalid)

    def test_image_cell_size_validation(self) -> None:
        self.assertEqual(validate_image_cell_size("17", "80"), (17.0, 80.0))
        self.assertEqual(validate_image_cell_size("22.5", "120"), (22.5, 120.0))
        for width, height in (("abc", "80"), ("17", "x"), (0, 80), (17, 0), (256, 80), (17, 410)):
            with self.subTest(width=width, height=height):
                with self.assertRaises(ValueError):
                    validate_image_cell_size(width, height)

    def test_custom_cell_size_changes_cell_and_image_display_only(self) -> None:
        custom_width = 24.0
        custom_height = 120.0
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "1-01-1.png"
            original_size = (640, 480)
            Image.new("RGB", original_size, "navy").save(source)
            output = root / "custom_size.xlsx"

            summary = create_excel(
                [source],
                output,
                image_column_width=custom_width,
                image_row_height=custom_height,
            )
            self.assertEqual(summary.inserted_count, 1)

            with zipfile.ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
                drawing = ElementTree.fromstring(archive.read("xl/drawings/drawing1.xml"))
                ns = {
                    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                }
                anchor = drawing.find("xdr:twoCellAnchor", ns)
                self.assertIsNotNone(anchor)
                start_col = int(anchor.findtext("xdr:from/xdr:col", namespaces=ns))
                start_row = int(anchor.findtext("xdr:from/xdr:row", namespaces=ns))
                end_col = int(anchor.findtext("xdr:to/xdr:col", namespaces=ns))
                end_row = int(anchor.findtext("xdr:to/xdr:row", namespaces=ns))
                end_col_offset = int(anchor.findtext("xdr:to/xdr:colOff", namespaces=ns))
                end_row_offset = int(anchor.findtext("xdr:to/xdr:rowOff", namespaces=ns))
                width_px = (
                    (end_col - start_col) * excel_column_width_to_pixels(custom_width)
                    + end_col_offset / 9525
                )
                height_px = (
                    (end_row - start_row) * excel_row_height_to_pixels(custom_height)
                    + end_row_offset / 9525
                )
                self.assertAlmostEqual(
                    width_px, excel_column_width_to_pixels(custom_width), places=6
                )
                self.assertAlmostEqual(
                    height_px, excel_row_height_to_pixels(custom_height), places=6
                )

                media_name = next(
                    name for name in archive.namelist() if name.startswith("xl/media/image")
                )
                with Image.open(BytesIO(archive.read(media_name))) as embedded:
                    self.assertEqual(embedded.size, original_size)

            workbook = load_workbook(output, data_only=False)
            sheet = workbook["이미지 정리"]
            self.assertAlmostEqual(
                sheet.row_dimensions[DATA_START_ROW + 1].height, custom_height
            )
            self.assertGreater(sheet.column_dimensions["B"].width, custom_width)
            self.assertLess(sheet.column_dimensions["B"].width, custom_width + 1)
            workbook.close()

    def test_natural_order_regression(self) -> None:
        paths = [Path("1_01_03.tif"), Path("1_1_02.tif"), Path("1_1_01.tif")]
        groups = organize_paths(paths)
        self.assertEqual(len(groups), 1)
        self.assertEqual(
            [item.path.name for item in groups[0][1]],
            ["1_1_01.tif", "1_1_02.tif", "1_01_03.tif"],
        )
        self.assertEqual(
            [item.sequence_text for item in groups[0][1]],
            ["1_01", "1_02", "01_03"],
        )

    def test_manual_order_only_changes_order_inside_same_group(self) -> None:
        paths = [
            Path("2_1_01.tif"),
            Path("1_1_01.tif"),
            Path("1_1_02.tif"),
            Path("2_1_02.tif"),
        ]
        override = [
            Path("1_1_02.tif"),
            Path("1_1_01.tif"),
            Path("2_1_02.tif"),
            Path("2_1_01.tif"),
        ]
        groups = organize_paths(paths, manual_order=override)
        self.assertEqual([name for name, _ in groups], ["1", "2"])
        self.assertEqual(
            [item.path.name for item in groups[0][1]],
            ["1_1_02.tif", "1_1_01.tif"],
        )
        self.assertEqual(
            [item.path.name for item in groups[1][1]],
            ["2_1_02.tif", "2_1_01.tif"],
        )

    def test_preflight_duplicate_and_missing_warnings(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            folder_a = root / "a"
            folder_b = root / "b"
            folder_a.mkdir()
            folder_b.mkdir()
            same_a = folder_a / "1_1_01.png"
            same_b = folder_b / "1_1_01.png"
            p3 = folder_a / "1_1_03.png"
            p4 = folder_a / "1_1_04.png"
            for path in (same_a, same_b, p3, p4):
                Image.new("RGB", (10, 10), "black").save(path)
            warnings = analyze_selection([same_a, same_b, p3, p4], check_missing=True)
            self.assertTrue(any("동일 파일명" in warning for warning in warnings))
            self.assertTrue(any("번호 누락 가능성" in warning for warning in warnings))

    def test_settings_round_trip(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "settings.json"
            expected = {
                "remember_settings": True,
                "image_column_width": 24.5,
                "image_row_height": 120,
                "check_missing": True,
            }
            save_settings(expected, path)
            self.assertEqual(load_settings(path), expected)
            path.write_text("{broken", encoding="utf-8")
            self.assertEqual(load_settings(path), {})

    def test_cancel_stops_before_final_output_is_published(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            paths = []
            for index in range(3):
                path = root / f"1_1_0{index + 1}.png"
                Image.new("RGB", (200, 120), "blue").save(path)
                paths.append(path)
            output = root / "cancelled.xlsx"
            calls = 0

            def cancel_check() -> bool:
                nonlocal calls
                calls += 1
                return calls >= 2

            with self.assertRaises(GenerationCancelled):
                create_excel(paths, output, cancel_check=cancel_check)
            self.assertFalse(output.exists())
            self.assertEqual(list(root.glob(".*.tmp.xlsx")), [])

    def test_no_groups_are_naturally_ordered(self) -> None:
        paths = [
            Path("No10-1.png"),
            Path("No2-3.png"),
            Path("No1-2.png"),
            Path("No2-1.png"),
            Path("No1-1.png"),
        ]
        groups = organize_paths(paths)
        self.assertEqual([name for name, _ in groups], ["No1", "No2", "No10"])
        self.assertEqual(
            [item.path.name for item in groups[1][1]],
            ["No2-1.png", "No2-3.png"],
        )

    def test_mixed_separators_do_not_override_numeric_sequence(self) -> None:
        paths = [Path("7-1-10.png"), Path("7_1_2.png"), Path("7 1 3.png")]
        groups = organize_paths(paths)
        self.assertEqual([name for name, _ in groups], ["7"])
        self.assertEqual(
            [item.path.name for item in groups[0][1]],
            ["7_1_2.png", "7 1 3.png", "7-1-10.png"],
        )
        self.assertLess(sequence_sort_key("1_2"), sequence_sort_key("1-10"))

    def test_generalized_prefix_groups_and_digit_lengths(self) -> None:
        paths = [
            Path("Lot10-2.png"),
            Path("No2_item_1.png"),
            Path("Lot2_10.png"),
            Path("No1-test-10.png"),
            Path("Lot2-2.png"),
            Path("No1_test_2.png"),
        ]
        groups = organize_paths(paths)
        self.assertEqual(
            {group: [item.path.name for item in items] for group, items in groups},
            {
                "Lot2": ["Lot2-2.png", "Lot2_10.png"],
                "Lot10": ["Lot10-2.png"],
                "No1": ["No1_test_2.png", "No1-test-10.png"],
                "No2": ["No2_item_1.png"],
            },
        )

    def test_examples_group_by_row_and_generate_valid_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            specs = [
                ("2-02-2.webp", "WEBP"),
                ("1-02-2.bmp", "BMP"),
                ("1-01-2.jpeg", "JPEG"),
                ("2-01-1.gif", "GIF"),
                ("1-02-1.png", "PNG"),
                ("2-02-1.tiff", "TIFF"),
                ("1-01-1.tif", "TIFF"),
                ("2-01-2.jpg", "JPEG"),
            ]

            paths: list[Path] = []
            for index, (name, image_format) in enumerate(specs):
                path = root / name
                image = Image.new("RGB", (240 + index, 180 + index), (20 * index, 80, 180))
                draw = ImageDraw.Draw(image)
                draw.text((10, 10), name, fill=(255, 255, 255))
                image.save(path, format=image_format)
                image.close()
                paths.append(path)

            output = root / "result.xlsx"
            summary = create_excel(paths, output)

            self.assertTrue(output.is_file())
            self.assertEqual(summary.selected_count, 8)
            self.assertEqual(summary.inserted_count, 8)
            self.assertEqual(summary.failed_count, 0)
            self.assertEqual(summary.group_count, 2)

            # ZIP CRC와 이미지 개체 수 검사.
            with zipfile.ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
                drawing = ElementTree.fromstring(archive.read("xl/drawings/drawing1.xml"))
                blips = list(
                    drawing.iter(
                        "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
                    )
                )
                self.assertEqual(len(blips), 8)

                ns = {
                    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                }
                first_anchor = drawing.find("xdr:twoCellAnchor", ns)
                self.assertIsNotNone(first_anchor)
                start_col = int(first_anchor.findtext("xdr:from/xdr:col", namespaces=ns))
                start_row = int(first_anchor.findtext("xdr:from/xdr:row", namespaces=ns))
                end_col = int(first_anchor.findtext("xdr:to/xdr:col", namespaces=ns))
                end_row = int(first_anchor.findtext("xdr:to/xdr:row", namespaces=ns))
                end_col_offset = int(
                    first_anchor.findtext("xdr:to/xdr:colOff", namespaces=ns)
                )
                end_row_offset = int(
                    first_anchor.findtext("xdr:to/xdr:rowOff", namespaces=ns)
                )
                # 첫 이미지가 이미지 표의 행 그룹 열(A) 다음인 B열에서 시작해야 한다.
                self.assertEqual(start_col, 1)
                # OOXML은 셀 경계를 (다음 셀, offset 0) 또는
                # (현재 셀, offset=셀 전체 크기)로 동등하게 표현할 수 있다.
                width_px = (
                    (end_col - start_col) * excel_column_width_to_pixels(IMAGE_COLUMN_WIDTH)
                    + end_col_offset / 9525
                )
                height_px = (
                    (end_row - start_row) * excel_row_height_to_pixels(IMAGE_ROW_HEIGHT)
                    + end_row_offset / 9525
                )
                self.assertAlmostEqual(
                    width_px, excel_column_width_to_pixels(IMAGE_COLUMN_WIDTH), places=6
                )
                self.assertAlmostEqual(
                    height_px, excel_row_height_to_pixels(IMAGE_ROW_HEIGHT), places=6
                )
                self.assertEqual(excel_row_height_to_pixels(IMAGE_ROW_HEIGHT), 106)

            # 제3자 파서(openpyxl)로 다시 열어 구조/크기/검증표 순서를 확인.
            workbook = load_workbook(output, data_only=False)
            sheet = workbook["이미지 정리"]
            self.assertEqual(len(sheet._images), 8)
            self.assertAlmostEqual(sheet.row_dimensions[DATA_START_ROW + 1].height, IMAGE_ROW_HEIGHT)

            # 이미지 정렬표 첫 열(A)은 행 그룹, 실제 이미지는 B열부터 시작한다.
            self.assertEqual(sheet["A1"].value, "이미지 정렬표")
            self.assertEqual(sheet["A2"].value, "행 그룹")
            self.assertEqual(sheet["B2"].value, "이미지 1")
            self.assertEqual(sheet["A3"].value, "1")
            self.assertEqual(sheet["A4"].value, "2")

            # XlsxWriter는 Excel UI 너비 17을 XML 내부 너비로 환산해 저장한다.
            self.assertGreater(sheet.column_dimensions["B"].width, IMAGE_COLUMN_WIDTH)
            self.assertLess(sheet.column_dimensions["B"].width, IMAGE_COLUMN_WIDTH + 1)

            validation_start_excel_col = 4 + 3  # 행 그룹 1 + 이미지 4 + 빈 열 1 다음(1 기반)
            self.assertEqual(sheet.cell(3, validation_start_excel_col).value, "1")
            self.assertEqual(
                [sheet.cell(3, validation_start_excel_col + i).value for i in range(1, 5)],
                ["1-01-1.tif", "1-01-2.jpeg", "1-02-1.png", "1-02-2.bmp"],
            )
            self.assertEqual(sheet.cell(4, validation_start_excel_col).value, "2")
            self.assertEqual(
                [sheet.cell(4, validation_start_excel_col + i).value for i in range(1, 5)],
                ["2-01-1.gif", "2-01-2.jpg", "2-02-1.tiff", "2-02-2.webp"],
            )
            workbook.close()

    def test_unknown_extension_is_read_by_file_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            disguised = root / "No1-1.unknown_extension"
            image = Image.new("RGB", (100, 100), "navy")
            image.save(disguised, format="PNG")
            image.close()

            output = root / "unknown.xlsx"
            summary = create_excel([disguised], output)
            self.assertEqual(summary.inserted_count, 1)
            self.assertEqual(summary.failed_count, 0)

    def test_misleading_svg_extension_uses_actual_raster_content(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            disguised = root / "No1-1.svg"
            image = Image.new("RGB", (90, 70), "purple")
            image.save(disguised, format="PNG")
            image.close()

            output = root / "misleading_extension.xlsx"
            summary = create_excel([disguised], output)
            self.assertEqual(summary.inserted_count, 1)
            self.assertEqual(summary.results[0].source_format, "PNG")

    def test_svg_content_detection_does_not_depend_on_extension(self) -> None:
        from excel_image_organizer_ux import _read_svg_source

        svg = (
            b'<svg xmlns="http://www.w3.org/2000/svg" width="20" height="10">'
            b'<rect width="20" height="10" fill="red"/></svg>'
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            plain = root / "No1-1.image_data"
            compressed = root / "No1-2.another_data"
            plain.write_bytes(svg)
            compressed.write_bytes(gzip.compress(svg))
            self.assertEqual(_read_svg_source(plain), svg)
            self.assertEqual(_read_svg_source(compressed), svg)

            truncated_gzip = root / "No1-3.broken_data"
            truncated_gzip.write_bytes(b"\x1f\x8b\x08")
            self.assertIsNone(_read_svg_source(truncated_gzip))

    def test_bad_file_keeps_position_and_records_failure(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            valid = root / "1-01-1.png"
            invalid = root / "1-01-2.tif"
            Image.new("RGB", (80, 60), "green").save(valid)
            invalid.write_text("not an image", encoding="utf-8")

            output = root / "partial.xlsx"
            summary = create_excel([invalid, valid], output)
            self.assertEqual(summary.selected_count, 2)
            self.assertEqual(summary.inserted_count, 1)
            self.assertEqual(summary.failed_count, 1)

            workbook = load_workbook(output)
            sheet = workbook["이미지 정리"]
            self.assertEqual(len(sheet._images), 1)
            # A열은 행 그룹이므로 두 번째 이미지 위치는 C열이다.
            self.assertIn("삽입 실패", sheet.cell(DATA_START_ROW + 1, 3).value)
            result_sheet = workbook["처리 결과"]
            self.assertEqual(result_sheet.cell(3, 9).value, "실패")
            workbook.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)