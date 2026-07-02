"""
pptx_chart_editor.py
====================
Đọc và chỉnh sửa data của các chart trong file .pptx trực tiếp qua XML cache.
Hoạt động với cả chart embedded lẫn chart đã được fix external link.

Cài đặt:
    pip install lxml openpyxl

Sử dụng nhanh:
    editor = PptxChartEditor("Template_fixed.pptx")
    editor.list_charts()
    editor.update_chart("chart7.xml", series_index=0, values=[100, 200, 300, 400, 500, 600, 700])
    editor.save("Template_updated.pptx")
"""

import os
import shutil
import zipfile
import tempfile
from copy import deepcopy
from lxml import etree

# XML namespaces
C  = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class PptxChartEditor:
    def __init__(self, pptx_path: str):
        """
        Mở file PPTX và đọc tất cả chart vào bộ nhớ.
        
        Args:
            pptx_path: Đường dẫn tới file .pptx
        """
        self.pptx_path = pptx_path
        self.tmp_dir = tempfile.mkdtemp(prefix="pptx_edit_")
        
        # Giải nén pptx vào thư mục tạm
        shutil.unpack_archive(pptx_path, self.tmp_dir, "zip")
        
        self.charts_dir = os.path.join(self.tmp_dir, "ppt", "charts")
        self._chart_trees = {}  
        print(f"✅ Loaded: {pptx_path}")
        print(f"   Charts found: {len(self._get_chart_files())}")


    def list_charts(self) -> list[dict]:
        """
        In ra danh sách tất cả chart trong file kèm thông tin series và data.
        
        Returns:
            List of chart info dicts
        """
        charts = []
        for chart_file in self._get_chart_files():
            info = self._parse_chart_info(chart_file)
            charts.append(info)
            print(f"\n📊 {chart_file}  [{info['type']}]  — {len(info['series'])} series")
            for i, s in enumerate(info['series']):
                cats_preview = s['categories'][:3]
                vals_preview = s['values'][:3]
                more = f" ... (+{len(s['values'])-3} more)" if len(s['values']) > 3 else ""
                print(f"   [{i}] '{s['name']}'")
                n_cats = len(s['categories'])
                cat_suffix = "" if n_cats <= 3 else f" ... (+{n_cats-3} more)"
                print(f"       categories : {cats_preview}{cat_suffix}")
                print(f"       values     : {vals_preview}{more}")
        return charts

    def get_chart_info(self, chart_file: str) -> dict:
        """
        Trả về thông tin chi tiết của một chart.
        
        Args:
            chart_file: Tên file chart, ví dụ "chart7.xml"
        
        Returns:
            Dict chứa type, series, categories, values
        """
        return self._parse_chart_info(chart_file)

    def update_chart(
        self,
        chart_file: str,
        series_index: int,
        values: list = None,
        categories: list = None,
        series_name: str = None,
    ):
        """
        Cập nhật data của một series trong chart.
        
        Args:
            chart_file:   Tên file chart, ví dụ "chart7.xml"
            series_index: Index của series cần sửa (bắt đầu từ 0)
            values:       List giá trị mới (số). None = giữ nguyên
            categories:   List category/label mới (string). None = giữ nguyên
            series_name:  Tên series mới. None = giữ nguyên
        
        Ví dụ:
            editor.update_chart("chart7.xml", series_index=0,
                                values=[100, 200, 150, 300, 250, 400, 350])
            
            editor.update_chart("chart8.xml", series_index=1,
                                categories=["MSN", "Vingroup", "Hòa Phát", "MWG", "Vinamilk"],
                                values=[3000, 28000, 7000, 6500, 5000],
                                series_name="Tuần mới")
        """
        root = self._get_tree(chart_file).getroot()
        series = root.findall(f".//{{{C}}}ser")
        
        if series_index >= len(series):
            raise IndexError(f"Chart '{chart_file}' chỉ có {len(series)} series (index 0..{len(series)-1})")
        
        ser = series[series_index]
        
        if series_name is not None:
            self._set_series_name(ser, series_name)
        
        if categories is not None:
            self._clear_pts(ser, f".//{{{C}}}cat")
            self._set_str_values(ser, f".//{{{C}}}cat", categories)
            self._update_count(ser, f".//{{{C}}}cat", len(categories))
        
        if values is not None:
            # values có thể nằm trong <val> hoặc <yVal> (scatter)
            val_tag = f".//{{{C}}}val" if ser.find(f".//{{{C}}}val") is not None else f".//{{{C}}}yVal"
            self._clear_pts(ser, val_tag)
            self._set_num_values(ser, val_tag, values)
            self._update_count(ser, val_tag, len(values))
        
        print(f"✅ Updated chart='{chart_file}' series[{series_index}]"
              + (f"  values={values[:3]}{'...' if values and len(values)>3 else ''}" if values else "")
              + (f"  categories={categories[:3]}..." if categories and len(categories)>3 else f"  categories={categories}" if categories else "")
              + (f"  name='{series_name}'" if series_name else ""))

    def update_all_series(self, chart_file: str, data: list[dict]):
        """
        Cập nhật nhiều series cùng lúc.
        
        Args:
            chart_file: Tên file chart
            data: List các dict, mỗi dict tương ứng 1 series:
                  [
                    {"series_index": 0, "values": [...], "categories": [...], "series_name": "..."},
                    {"series_index": 1, "values": [...]},
                    ...
                  ]
        
        Ví dụ:
            editor.update_all_series("chart4.xml", [
                {"series_index": 0, "values": [0.10, 0.15, 0.12, 0.20, 0.09]},
                {"series_index": 1, "values": [0.80, 0.75, 0.82, 0.70, 0.85]},
                {"series_index": 2, "values": [0.10, 0.10, 0.06, 0.10, 0.06]},
            ])
        """
        root = self._get_tree(chart_file).getroot()
        series_list = root.findall(f".//{{{C}}}ser")

        for item in data:
            idx = item["series_index"]
            # Xóa toàn bộ <c:pt> cũ trước khi ghi mới
            if idx < len(series_list):
                ser = series_list[idx]
                for xpath in [f".//{{{C}}}cat", f".//{{{C}}}val", f".//{{{C}}}yVal"]:
                    container = ser.find(xpath)
                    if container is None:
                        continue
                    for cache_tag in ["strCache", "numCache"]:
                        cache = container.find(f"{{{C}}}{cache_tag}")
                        if cache is not None:
                            for pt in cache.findall(f"{{{C}}}pt"):
                                cache.remove(pt)

            self.update_chart(
                chart_file,
                series_index=idx,
                values=item.get("values"),
                categories=item.get("categories"),
                series_name=item.get("series_name"),
            )

    def clear_series_beyond(self, chart_file: str, keep_n: int) -> None:
        """
        Clear series name + data pts for all series with index >= keep_n.

        Used when the template has more series slots than the actual data
        (e.g. chart7 has 5 brand slots but only 3 brands are provided).
        Blanks series name and zeroes ptCount so PowerPoint renders nothing
        for those series.  Also injects <c:legendEntry delete="1"/> for each
        hidden index so PowerPoint omits them from the legend entirely.
        """
        root = self._get_tree(chart_file).getroot()
        all_ser = root.findall(f".//{{{C}}}ser")
        total_series = len(all_ser)
        cleared = 0
        for ser in all_ser[keep_n:]:
            # Blank series name
            v_el = ser.find(f".//{{{C}}}tx//{{{C}}}v")
            if v_el is not None:
                v_el.text = ""
            # Clear data points
            for container_tag in (f"{{{C}}}cat", f"{{{C}}}val", f"{{{C}}}yVal"):
                container = ser.find(f".//{container_tag}")
                if container is None:
                    continue
                for cache_tag in (f"{{{C}}}strCache", f"{{{C}}}numCache"):
                    cache = container.find(f".//{cache_tag}")
                    if cache is None:
                        continue
                    for pt in cache.findall(f"{{{C}}}pt"):
                        cache.remove(pt)
                    pt_count = cache.find(f"{{{C}}}ptCount")
                    if pt_count is not None:
                        pt_count.set("val", "0")
            cleared += 1

        # ── Sync legend entries ────────────────────────────────────────────────
        # Remove all existing legendEntry elements first (idempotent on re-runs)
        legend = root.find(f".//{{{C}}}legend")
        if legend is not None:
            for old in legend.findall(f"{{{C}}}legendEntry"):
                legend.remove(old)
            # Insert <c:legendEntry><c:idx val="N"/><c:delete val="1"/></c:legendEntry>
            # before the first non-legendEntry child so order is OOXML-compliant.
            insert_pos = 0
            for hidden_idx in range(keep_n, total_series):
                entry = etree.Element(f"{{{C}}}legendEntry")
                idx_el = etree.SubElement(entry, f"{{{C}}}idx")
                idx_el.set("val", str(hidden_idx))
                del_el = etree.SubElement(entry, f"{{{C}}}delete")
                del_el.set("val", "1")
                legend.insert(insert_pos, entry)
                insert_pos += 1

        if cleared:
            print(f"✅ Cleared {cleared} excess series in '{chart_file}' (keeping {keep_n})")

    def save(self, output_path: str = None):
        """
        Lưu tất cả thay đổi vào file .pptx mới.
        
        Args:
            output_path: Đường dẫn file output. 
                         Mặc định: thêm '_updated' vào tên file gốc.
        """
        if output_path is None:
            base, ext = os.path.splitext(self.pptx_path)
            output_path = f"{base}_updated{ext}"
        
        # Ghi tất cả XML tree đang cached xuống disk
        for chart_file, tree in self._chart_trees.items():
            path = os.path.join(self.charts_dir, chart_file)
            tree.write(path, xml_declaration=True, encoding="UTF-8", standalone=True)
        
        # Repack thành pptx
        self._pack_pptx(self.tmp_dir, output_path)
        print(f"\n💾 Saved → {output_path}")
        return output_path

    def __del__(self):
        """Dọn thư mục tạm khi object bị hủy."""
        try:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
        except Exception:
            pass

    # ─────────────────────────────────────────────
    # INTERNAL HELPERS
    # ─────────────────────────────────────────────

    def _get_chart_files(self) -> list[str]:
        if not os.path.exists(self.charts_dir):
            return []
        files = [f for f in os.listdir(self.charts_dir)
                 if f.startswith("chart") and f.endswith(".xml")]
        return sorted(files, key=lambda x: int(x.replace("chart","").replace(".xml","")))

    def _get_tree(self, chart_file: str):
        if chart_file not in self._chart_trees:
            path = os.path.join(self.charts_dir, chart_file)
            if not os.path.exists(path):
                raise FileNotFoundError(f"Không tìm thấy: {chart_file}")
            self._chart_trees[chart_file] = etree.parse(path)
        return self._chart_trees[chart_file]

    def _parse_chart_info(self, chart_file: str) -> dict:
        root = self._get_tree(chart_file).getroot()
        
        # Chart type
        chart_types = []
        for tag in ["barChart","lineChart","pieChart","doughnutChart",
                    "areaChart","scatterChart","radarChart","bubbleChart"]:
            if root.findall(f".//{{{C}}}{tag}"):
                chart_types.append(tag.replace("Chart",""))
        
        series_list = []
        for ser in root.findall(f".//{{{C}}}ser"):
            name_v = ser.find(f".//{{{C}}}tx//{{{C}}}v")
            name = name_v.text if name_v is not None else f"Series"
            cats = [v.text for v in ser.findall(f".//{{{C}}}cat//{{{C}}}v")]
            vals = [v.text for v in ser.findall(f".//{{{C}}}val//{{{C}}}v")]
            if not vals:
                vals = [v.text for v in ser.findall(f".//{{{C}}}yVal//{{{C}}}v")]
            # Convert numeric strings
            vals_out = []
            for v in vals:
                try:
                    vals_out.append(float(v))
                except (TypeError, ValueError):
                    vals_out.append(v)
            series_list.append({"name": name, "categories": cats, "values": vals_out})
        
        return {"chart": chart_file, "type": "/".join(chart_types), "series": series_list}

    def _clear_pts(self, ser, xpath: str):
        """Xóa toàn bộ <c:pt> cũ trong container (cat hoặc val)."""
        container = ser.find(xpath)
        if container is None:
            return
        for cache_tag in ["strCache", "numCache"]:
            cache = container.find(f"{{{C}}}{cache_tag}")
            if cache is not None:
                for pt in cache.findall(f"{{{C}}}pt"):
                    cache.remove(pt)

    def _set_series_name(self, ser, new_name: str):
        v_el = ser.find(f".//{{{C}}}tx//{{{C}}}v")
        if v_el is not None:
            v_el.text = new_name

    def _set_str_values(self, ser, xpath: str, values: list):
        """Ghi list string vào <c:v> bên trong xpath (dùng cho categories)."""
        container = ser.find(xpath)
        if container is None:
            return
        v_elements = container.findall(f".//{{{C}}}v")
        
        # Nếu số phần tử khác nhau → rebuild
        if len(v_elements) != len(values):
            # Tìm parent của các <v> (thường là <c:strCache> hoặc <c:numCache>)
            cache = container.find(f".//{{{C}}}strCache")
            if cache is None:
                cache = container.find(f".//{{{C}}}numCache")
            if cache is not None:
                for old_v in cache.findall(f"{{{C}}}pt"):
                    cache.remove(old_v)
                for i, val in enumerate(values):
                    pt = etree.SubElement(cache, f"{{{C}}}pt")
                    pt.set("idx", str(i))
                    v = etree.SubElement(pt, f"{{{C}}}v")
                    v.text = str(val)
                return
        
        for v_el, val in zip(v_elements, values):
            v_el.text = str(val)

    def _set_num_values(self, ser, xpath: str, values: list):
        """Ghi list số vào <c:v> bên trong xpath (dùng cho values)."""
        container = ser.find(xpath)
        if container is None:
            return
        v_elements = container.findall(f".//{{{C}}}v")
        
        if len(v_elements) != len(values):
            cache = container.find(f".//{{{C}}}numCache")
            if cache is not None:
                for old_pt in cache.findall(f"{{{C}}}pt"):
                    cache.remove(old_pt)
                for i, val in enumerate(values):
                    pt = etree.SubElement(cache, f"{{{C}}}pt")
                    pt.set("idx", str(i))
                    v = etree.SubElement(pt, f"{{{C}}}v")
                    v.text = str(val)
                return
        
        for v_el, val in zip(v_elements, values):
            v_el.text = str(val)

    def _update_count(self, ser, xpath: str, count: int):
        """Cập nhật <c:ptCount val="..."> cho đúng số điểm data."""
        container = ser.find(xpath)
        if container is None:
            return
        for pt_count in container.findall(f".//{{{C}}}ptCount"):
            pt_count.set("val", str(count))

    @staticmethod
    def _pack_pptx(src_dir: str, output_path: str):
        """Nén thư mục thành file .pptx."""
        tmp_zip = output_path + ".tmp.zip"
        with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(src_dir):
                # Loại bỏ __MACOSX và .DS_Store
                dirs[:] = [d for d in dirs if d != "__MACOSX"]
                for file in files:
                    if file == ".DS_Store":
                        continue
                    full_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(full_path, src_dir)
                    zf.write(full_path, arcname)
        os.replace(tmp_zip, output_path)


