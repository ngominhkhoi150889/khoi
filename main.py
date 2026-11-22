from datetime import date
import hashlib
import io
import json
import gspread
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quan ly Chi tieu", page_icon="💸", layout="wide"
)


@st.cache_resource
def init_gsheet():
  raw = st.secrets.get("gcp_service_account")
  creds = json.loads(raw) if isinstance(raw, str) else dict(raw)
  client = gspread.service_account_from_dict(creds)
  client.http_client.set_timeout(15)
  sh = client.open_by_key(st.secrets.get("sheet_id"))
  return {
      "ct": sh.worksheet("chi_tieu"),
      "us": sh.worksheet("users"),
      "cf": sh.worksheet("config"),
  }


try:
  sheets = init_gsheet()
  ws_ct = sheets.get("ct")
  ws_us = sheets.get("us")
  ws_cf = sheets.get("cf")
except Exception as e:
  st.error(f"Loi ket noi Google Sheets: {e}")
  st.stop()


def hash_pw(pw):
  return hashlib.sha256(pw.encode()).hexdigest()


@st.cache_data(ttl=120)
def load_users():
  rows = ws_us.get_all_values()
  users = {}
  for r in rows[1:]:
    if len(r) >= 2 and str(r).strip() != "":
      u = str(r).strip()
      dname = str(r).strip() if len(r) >= 3 and str(r).strip() else u
      rl = str(r).strip() if len(r) >= 4 and str(r).strip() else "member"
      users[u] = {"password": str(r).strip(), "display_name": dname, "role": rl}
  if "admin" not in users:
    apw = hash_pw("SUPERACCOUNT@123")
    ws_us.append_row(["admin", apw, "Admin", "admin"])
    users["admin"] = {"password": apw, "display_name": "Admin", "role": "admin"}
  return users


def add_user(u, pw, dname, role="member"):
  ws_us.append_row([u, hash_pw(pw), dname, role])
  load_users.clear()


def del_user(u):
  cell = ws_us.find(u)
  if cell:
    ws_us.delete_rows(cell.row)
    load_users.clear()


@st.cache_data(ttl=60)
def load_data():
  rows = ws_ct.get_all_values()
  cols = ["Ngay", "Danh_muc", "Mo_ta", "So_tien", "Nguoi_tao"]
  data = []
  for r in rows[1:]:
    if len(r) >= 4 and str(r).strip() != "":
      r5 = r[:5] if len(r) >= 5 else r + ["Admin"]
      data.append(r5)
  if not data:
    return pd.DataFrame(columns=cols)
  df = pd.DataFrame(data, columns=cols)
  df["Ngay"] = pd.to_datetime(df["Ngay"], errors="coerce").dt.date
  df = df.dropna(subset=["Ngay"])
  df["So_tien"] = (
      df["So_tien"].astype(str).str.replace(",", "").str.replace(".", "")
  )
  df["So_tien"] = pd.to_numeric(df["So_tien"], errors="coerce").fillna(0)
  return df


def save_all(df):
  ws_ct.clear()
  ws_ct.append_row(["Ngày", "Danh mục", "Mô tả", "Số tiền", "Người tạo"])
  if not df.empty:
    dfs = df.copy()
    dfs["Ngay"] = dfs["Ngay"].astype(str)
    ws_ct.append_rows(dfs.values.tolist())
  load_data.clear()


@st.cache_data(ttl=120)
def load_budget():
  rows = ws_cf.get_all_values()
  for r in rows[1:]:
    if len(r) >= 2 and str(r).strip() == "han_muc":
      try:
        return int(float(str(r).strip()))
      except ValueError:
        return 5000000
  ws_cf.append_row(["han_muc", 5000000])
  return 5000000


def save_budget(amt):
  cell = ws_cf.find("han_muc")
  if cell:
    ws_cf.update_cell(cell.row, 2, int(amt))
  else:
    ws_cf.append_row(["han_muc", int(amt)])
  load_budget.clear()


def to_excel(df_ex):
  out = io.BytesIO()
  df_out = df_ex.rename(
      columns={
          "Ngay": "Ngày",
          "Danh_muc": "Danh mục",
          "Mo_ta": "Mô tả",
          "So_tien": "Số tiền",
          "Nguoi_tao": "Người tạo",
      }
  )
  with pd.ExcelWriter(out, engine="openpyxl") as writer:
    df_out.to_excel(writer, index=False, sheet_name="ChiTieu")
  return out.getvalue()


if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.username = ""
  st.session_state.display_name = ""
  st.session_state.is_admin = False

users_db = load_users()

if not st.session_state.logged_in:
  st.title("🔐 Đăng nhập Hệ thống Quản lý Chi tiêu")
  c_l, c_m, c_r = st.columns([1, 1.5, 1])
  with c_m:
    t1, t2 = st.tabs(["Đăng nhập", "Tạo tài khoản mới"])
    with t1:
      with st.form("login_form"):
        ui = st.text_input("Tên đăng nhập").strip()
        pi = st.text_input("Mật khẩu", type="password")
        if st.form_submit_button("Đăng nhập", use_container_width=True):
          uinfo = users_db.get(ui)
          if uinfo and uinfo.get("password") == hash_pw(pi):
            st.session_state.logged_in = True
            st.session_state.username = ui
            st.session_state.display_name = uinfo.get("display_name")
            st.session_state.is_admin = (
                ui == "admin" or uinfo.get("role") == "admin"
            )
            st.rerun()
          else:
            st.error("Sai tên đăng nhập hoặc mật khẩu!")
    with t2:
      with st.form("reg_form"):
        nu = st.text_input("Tên đăng nhập mới (viết liền, không dấu)").strip()
        nd = st.text_input("Tên hiển thị (VD: Minh Khôi)").strip()
        npw = st.text_input("Mật khẩu", type="password")
        cpw = st.text_input("Nhập lại mật khẩu", type="password")
        if st.form_submit_button("Tạo tài khoản", use_container_width=True):
          if not nu or not nd or not npw:
            st.warning("Vui lòng điền đầy đủ thông tin.")
          elif nu.lower() == "admin" or nu in users_db:
            st.error("Tên đăng nhập đã tồn tại!")
          elif npw != cpw:
            st.error("Mật khẩu không khớp!")
          else:
            add_user(nu, npw, nd, "member")
            st.success(f"Đã tạo tài khoản **{nd}**! Chuyển sang tab Đăng nhập.")
  st.stop()

df = load_data()
han_muc = load_budget()
is_admin = st.session_state.is_admin

top1, top2 = st.columns([4, 1.2])
with top1:
  st.title("💸 Quản lý & Thống kê Chi tiêu")
with top2:
  badge = "👑 Quản trị viên" if is_admin else "👤 Thành viên"
  st.markdown(f"**{st.session_state.display_name}** ({badge})")
  if st.button("🚪 Đăng xuất", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()

hom_nay = date.today()
thang_now = hom_nay.strftime("%Y-%m")

cac_thang = {thang_now}
if not df.empty:
  cac_thang.update(pd.to_datetime(df["Ngay"]).dt.strftime("%Y-%m").unique())
ds_thang = sorted(list(cac_thang), reverse=True)

cf1, cf2, cf3 = st.columns([1.2, 1.2, 2])
with cf1:
  thang_chon = st.selectbox(
      "📅 Chọn lịch tháng:",
      ds_thang,
      format_func=lambda x: (
          f"Tháng {x.split('-')}/{x.split('-')}"
          + (" (Hiện tại)" if x == thang_now else " 🔒")
      ),
  )

with cf2:
  ds_nguoi = ["Tất cả"] + (
      sorted(df["Nguoi_tao"].unique().tolist()) if not df.empty else []
  )
  nguoi_chon = st.selectbox("👥 Lọc theo người tạo:", ds_nguoi)

bi_khoa = thang_chon < thang_now

with cf3:
  st.write("")
  st.write("")
  if bi_khoa:
    st.warning(
        f"🔒 **Tháng {thang_chon.split('-')}/{thang_chon.split('-')} đã khóa.** Không thể thêm hoặc xóa."
    )
  else:
    st.success(
        f"🔓 **Đang mở sổ tháng {thang_chon.split('-')}/{thang_chon.split('-')}.** Dữ liệu lưu Google Sheets."
    )

with st.sidebar:
  st.header("🎯 Hạn mức chi tiêu")
  if is_admin:
    hm_moi = st.number_input(
        "Ngân sách tháng (VNĐ)",
        min_value=100000,
        step=500000,
        value=int(han_muc),
    )
    if hm_moi != han_muc:
      save_budget(hm_moi)
      han_muc = hm_moi
      st.toast("Đã lưu hạn mức!")
  else:
    st.info(
        f"Ngân sách tháng: **{han_muc:,.0f} đ**\n*(Chỉ Admin được thay đổi)*"
    )

  st.divider()
  st.header("➕ Thêm khoản chi")
  if bi_khoa:
    st.error("Tháng này đã bị khóa sổ.")
  else:
    dau_thang = hom_nay.replace(day=1)
    with st.form("form_chi_tieu", clear_on_submit=True):
      ngay = st.date_input(
          "Ngày chi", value=hom_nay, min_value=dau_thang, max_value=hom_nay
      )
      dm = st.selectbox(
          "Danh mục",
          [
              "Ăn uống",
              "Di chuyển",
              "Mua sắm",
              "Hóa đơn & Tiện ích",
              "Giải trí",
              "Khác",
          ],
      )
      mota = st.text_input("Mô tả (VD: Cà phê, Đổ xăng...)")
      tien = st.number_input("Số tiền (VNĐ)", min_value=1000, step=5000)
      nguoi = st.session_state.display_name
      st.caption(f"Người tạo: **{nguoi}**")
      if st.form_submit_button("Lưu khoản chi", use_container_width=True):
        ws_ct.append_row(
            [str(ngay), dm, mota if mota else dm, int(tien), nguoi]
        )
        load_data.clear()
        st.rerun()

  if is_admin:
    st.divider()
    with st.expander("⚙️ Quản trị tài khoản (Admin)", expanded=False):
      m_keys = [u for u in users_db.keys() if u != "admin"]
      if not m_keys:
        st.write("Chưa có tài khoản thành viên nào khác.")
      else:
        for uk in m_keys:
          ui_info = users_db.get(uk)
          uc1, uc2 = st.columns()
          uc1.write(f"• **{ui_info.get('display_name')}** (`{uk}`)")
          if uc2.button("Xóa", key=f"del_u_{uk}"):
            del_user(uk)
            st.toast(f"Đã xóa {uk}")
            st.rerun()

if not df.empty:
  df["YM"] = pd.to_datetime(df["Ngay"]).dt.strftime("%Y-%m")
  df_t = df[df["YM"] == thang_chon].copy()
  if nguoi_chon != "Tất cả":
    df_t = df_t[df_t["Nguoi_tao"] == nguoi_chon]
else:
  df_t = pd.DataFrame(
      columns=["Ngay", "Danh_muc", "Mo_ta", "So_tien", "Nguoi_tao"]
  )

st.divider()

tong = df_t["So_tien"].sum() if not df_t.empty else 0
con_lai = han_muc - tong
ty_le = tong / han_muc if han_muc > 0 else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Tổng tiêu đã lọc", f"{tong:,.0f} đ")
c2.metric("Hạn mức tháng", f"{han_muc:,.0f} đ")
c3.metric(
    "Ngân sách còn lại",
    f"{con_lai:,.0f} đ",
    delta=f"{con_lai:,.0f} đ",
    delta_color="normal",
)
c4.metric("Số giao dịch", f"{len(df_t)} lần")

st.progress(min(ty_le, 1.0))
if ty_le >= 1.0:
  st.error(
      f"🚨 **Cảnh báo vượt ngân sách!** Vượt hạn mức {abs(con_lai):,.0f} đ ({ty_le*100:.1f}%)."
  )
elif ty_le >= 0.8:
  st.warning(
      f"⚠️ **Sắp chạm trần ngân sách:** Đã dùng {ty_le*100:.1f}% hạn mức."
  )
else:

    st.info(f"✅ Trong tầm kiểm soát ({ty_le*100:.1f}% ngân sách).")
    st.divider()
    ch1, ch2 = st.columns(2)
with ch1:
    st.subheader("📈 So sánh chi tiêu giữa các tháng")
if not df.empty:
    st.bar_chart(df.groupby("YM")["So_tien"].sum().sort_index())
else:
    st.caption("Chưa đủ dữ liệu để so sánh.")
with ch2:
    st.subheader(
    f"📊 Phân bổ danh mục (Tháng {thang_chon.split('-')}/{thang_chon.split('-')})"
    )
if not df_t.empty:
    st.bar_chart(df_t.groupby("Danh_muc")["So_tien"].sum())
else:
    st.caption("Không có khoản chi nào trong tháng này.")
    st.divider()
    hc1, hc2 = st.columns()
with hc1:
    st.subheader("📋 Chi tiết giao dịch")
with hc2:
    if not df_t.empty:
        df_ex = df_t.drop(columns=["YM"], errors="ignore")
        st.download_button(
    "📥 Xuất Excel (.xlsx)",
    to_excel(df_ex),
    f"Chi_tieu_{thang_chon}.xlsx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
    )
if df_t.empty:
    st.info("Chưa có dữ liệu chi tiêu nào phù hợp.")
else:
    df_s = df_t.sort_values(by="Ngay", ascending=False)
if bi_khoa:
    cols = st.columns([1.2, 1.5, 2.5, 1.5, 1.5])
    cols.markdown("Ngày")
    cols.markdown("Danh mục")
    cols.markdown("Mô tả")
    cols.markdown("Số tiền")
    cols.markdown("Người tạo")
else:
    cols = st.columns([1.2, 1.5, 2.3, 1.5, 1.5, 0.7])
    cols.markdown("Ngày")
    cols.markdown("Danh mục")
    cols.markdown("Mô tả")
    cols.markdown("Số tiền")
    cols.markdown("Người tạo")
    cols.markdown("Xóa")
    st.markdown("---")
for idx, row in df_s.iterrows():
    if bi_khoa:
        rc = st.columns([1.2, 1.5, 2.5, 1.5, 1.5])
        rc.write(str(row["Ngay"]))
        rc.write(row["Danh_muc"])
        rc.write(row["Mo_ta"])
        rc.write(f"{row['So_tien']:,.0f} đ")
        rc.write(f"👤 {row['Nguoi_tao']}")
else:
        rc = st.columns([1.2, 1.5, 2.3, 1.5, 1.5, 0.7])
        rc.write(str(row["Ngay"]))
        rc.write(row["Danh_muc"])
        rc.write(row["Mo_ta"])
        rc.write(f"{row['So_tien']:,.0f} đ")
        rc.write(f"👤 {row['Nguoi_tao']}")
        can_del = is_admin or (row["Nguoi_tao"] == st.session_state.display_name)
if can_del:
    if rc.button("🗑️", key=f"del_{idx}", help="Xóa khoản chi này"):
        df_u = df.drop(index=idx).drop(columns=["YM"], errors="ignore")
        save_all(df_u)
        st.rerun()
    else:
        rc.write("—")

