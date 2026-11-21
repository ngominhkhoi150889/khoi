from datetime import date
import hashlib
import io
import json
import gspread
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quản lý Chi tiêu", page_icon="💸", layout="wide"
)
status_box = st.empty()
status_box.info("⏳ Đang đọc cấu hình Secrets...")

try:
  raw_creds = st.secrets["gcp_service_account"]
  creds_dict = (
      json.loads(raw_creds) if isinstance(raw_creds, str) else dict(raw_creds)
  )
  sheet_id = st.secrets["sheet_id"]
except Exception as e:
  status_box.error(f"❌ Lỗi định dạng Secrets: {e}")
  st.stop()

status_box.info("⏳ Đang kết nối tới Google Sheets...")
try:
  client = gspread.service_account_from_dict(creds_dict)
  client.http_client.set_timeout(15)  # Giới hạn tối đa 15 giây để không xoay vô tận
  sh = client.open_by_key(sheet_id)
  ws_chitieu = sh.worksheet("chi_tieu")
  ws_users = sh.worksheet("users")
  ws_config = sh.worksheet("config")
  status_box.empty()  # Kết nối xong thì xóa dòng thông báo
except Exception as e:
  status_box.error(f"❌ Lỗi kết nối Google Sheets: {repr(e)}")
  st.stop()



# --- KẾT NỐI GOOGLE SHEETS ---
@st.cache_resource
def init_gsheet():
  raw_creds = st.secrets["gcp_service_account"]
  creds_dict = (
      json.loads(raw_creds) if isinstance(raw_creds, str) else dict(raw_creds)
  )
  client = gspread.service_account_from_dict(creds_dict)
  return client.open_by_key(st.secrets["sheet_id"])



try:
  sh = init_gsheet()
  ws_chitieu = sh.worksheet("chi_tieu")
  ws_users = sh.worksheet("users")
  ws_config = sh.worksheet("config")
except Exception as e:
  st.error(f"Lỗi kết nối Google Sheets. Hãy kiểm tra lại Secrets và tên Sheet: {e}")
  st.stop()


# --- HÀM BẢO MẬT & QUẢN LÝ TÀI KHOẢN ---
def hash_password(password):
  return hashlib.sha256(password.encode()).hexdigest()


def load_users():
  records = ws_users.get_all_records()
  users = {}
  for r in records:
    u_name = str(r.get("username", "")).strip()
    if u_name:
      users[u_name] = {
          "password": str(r.get("password", "")),
          "display_name": str(r.get("display_name", u_name)),
          "role": str(r.get("role", "member")),
      }
  if "admin" not in users:
    admin_pw = hash_password("SUPERACCOUNT@123")
    ws_users.append_row(["admin", admin_pw, "Admin", "admin"])
    users["admin"] = {
        "password": admin_pw,
        "display_name": "Admin",
        "role": "admin",
    }
  return users


def add_user_to_sheet(username, password, display_name, role="member"):
  ws_users.append_row([username, hash_password(password), display_name, role])


def delete_user_from_sheet(username):
  cell = ws_users.find(username)
  if cell:
    ws_users.delete_rows(cell.row)


# --- HÀM XỬ LÝ DỮ LIỆU & NGÂN SÁCH ---
def load_data():
  records = ws_chitieu.get_all_records()
  if not records:
    return pd.DataFrame(
        columns=["Ngày", "Danh mục", "Mô tả", "Số tiền", "Người tạo"]
    )
  df = pd.DataFrame(records)
  df["Ngày"] = pd.to_datetime(df["Ngày"]).dt.date
  df["Số tiền"] = pd.to_numeric(df["Số tiền"], errors="coerce").fillna(0)
  return df


def save_all_data(df):
  ws_chitieu.clear()
  ws_chitieu.append_row(["Ngày", "Danh mục", "Mô tả", "Số tiền", "Người tạo"])
  if not df.empty:
    df_save = df.copy()
    df_save["Ngày"] = df_save["Ngày"].astype(str)
    ws_chitieu.append_rows(df_save.values.tolist())


def load_budget():
  records = ws_config.get_all_records()
  for r in records:
    if r.get("key") == "han_muc":
      return int(r.get("value", 5000000))
  ws_config.append_row(["han_muc", 5000000])
  return 5000000


def save_budget(amount):
  cell = ws_config.find("han_muc")
  if cell:
    ws_config.update_cell(cell.row, 2, int(amount))
  else:
    ws_config.append_row(["han_muc", int(amount)])


def to_excel(df_export):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_export.to_excel(writer, index=False, sheet_name="ChiTieu")
  return output.getvalue()


# --- KHỞI TẠO TRẠNG THÁI ĐĂNG NHẬP ---
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
  st.session_state.username = ""
  st.session_state.display_name = ""
  st.session_state.is_admin = False

users_db = load_users()

# ==========================================
# MÀN HÌNH ĐĂNG NHẬP / ĐĂNG KÝ TÀI KHOẢN
# ==========================================
if not st.session_state.logged_in:
  st.title("🔐 Đăng nhập Hệ thống Quản lý Chi tiêu")
  col_empty1, col_auth, col_empty2 = st.columns([1, 1.5, 1])

  with col_auth:
    tab_login, tab_register = st.tabs(["Đăng nhập", "Tạo tài khoản mới"])

    with tab_login:
      with st.form("login_form"):
        user_input = st.text_input("Tên đăng nhập").strip()
        pass_input = st.text_input("Mật khẩu", type="password")
        if st.form_submit_button("Đăng nhập", use_container_width=True):
          if user_input in users_db and users_db[user_input][
              "password"
          ] == hash_password(pass_input):
            st.session_state.logged_in = True
            st.session_state.username = user_input
            st.session_state.display_name = users_db[user_input]["display_name"]
            st.session_state.is_admin = (
                user_input == "admin"
                or users_db[user_input].get("role") == "admin"
            )
            st.rerun()
          else:
            st.error("Sai tên đăng nhập hoặc mật khẩu!")

    with tab_register:
      with st.form("register_form"):
        new_user = st.text_input(
            "Tên đăng nhập mới (viết liền, không dấu)"
        ).strip()
        new_display = st.text_input("Tên hiển thị (VD: Minh Khôi)").strip()
        new_pass = st.text_input("Mật khẩu", type="password")
        confirm_pass = st.text_input("Nhập lại mật khẩu", type="password")
        if st.form_submit_button("Tạo tài khoản", use_container_width=True):
          if not new_user or not new_display or not new_pass:
            st.warning("Vui lòng điền đầy đủ thông tin.")
          elif new_user.lower() == "admin" or new_user in users_db:
            st.error("Tên đăng nhập này đã tồn tại hoặc không hợp lệ!")
          elif new_pass != confirm_pass:
            st.error("Mật khẩu nhập lại không khớp!")
          else:
            add_user_to_sheet(new_user, new_pass, new_display, "member")
            st.success(
                f"Đã tạo tài khoản **{new_display}**! Hãy chuyển sang tab Đăng nhập."
            )
  st.stop()

# ==========================================
# GIAO DIỆN CHÍNH SAU KHI ĐĂNG NHẬP
# ==========================================
df = load_data()
han_muc_thang = load_budget()
is_admin = st.session_state.is_admin

top_title, top_user = st.columns([4, 1.2])
with top_title:
  st.title("💸 Quản lý & Thống kê Chi tiêu")
with top_user:
  role_badge = "👑 Quản trị viên" if is_admin else "👤 Thành viên"
  st.markdown(f"**{st.session_state.display_name}** ({role_badge})")
  if st.button("🚪 Đăng xuất", use_container_width=True):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.display_name = ""
    st.session_state.is_admin = False
    st.rerun()

# --- XỬ LÝ LỊCH & TRẠNG THÁI KHÓA THÁNG ---
hom_nay = date.today()
thang_hien_tai_str = hom_nay.strftime("%Y-%m")

cac_thang = set([thang_hien_tai_str])
if not df.empty:
  cac_thang.update(pd.to_datetime(df["Ngày"]).dt.strftime("%Y-%m").unique())
danh_sach_thang = sorted(list(cac_thang), reverse=True)

col_filter, col_user_filter, col_status = st.columns([1.2, 1.2, 2])
with col_filter:
  thang_duoc_chon = st.selectbox(
      "📅 Chọn lịch tháng:",
      danh_sach_thang,
      format_func=lambda x: (
          f"Tháng {x.split('-')[1]}/{x.split('-')[0]}"
          + (" (Hiện tại)" if x == thang_hien_tai_str else " 🔒")
      ),
  )

with col_user_filter:
  danh_sach_nguoi_tao = ["Tất cả"] + (
      sorted(df["Người tạo"].unique().tolist()) if not df.empty else []
  )
  nguoi_tao_chon = st.selectbox("👥 Lọc theo người tạo:", danh_sach_nguoi_tao)

bi_khoa = thang_duoc_chon < thang_hien_tai_str

with col_status:
  st.write("")
  st.write("")
  if bi_khoa:
    st.warning(
        f"🔒 **Tháng {thang_duoc_chon.split('-')[1]}/{thang_duoc_chon.split('-')[0]} đã khóa.** Không thể thêm hoặc xóa."
    )
  else:
    st.success(
        f"🔓 **Đang mở sổ tháng {thang_duoc_chon.split('-')[1]}/{thang_duoc_chon.split('-')[0]}.** Dữ liệu lưu trực tiếp Google Sheets."
    )

# --- SIDEBAR: HẠN MỨC, NHẬP KHOẢN CHI & QUẢN TRỊ ---
with st.sidebar:
  st.header("🎯 Hạn mức chi tiêu")
  if is_admin:
    han_muc_moi = st.number_input(
        "Ngân sách tháng (VNĐ)",
        min_value=100000,
        step=500000,
        value=int(han_muc_thang),
    )
    if han_muc_moi != han_muc_thang:
      save_budget(han_muc_moi)
      han_muc_thang = han_muc_moi
      st.toast("Đã lưu hạn mức lên Google Sheets!")
  else:
    st.info(
        f"Ngân sách tháng: **{han_muc_thang:,.0f} đ**\n*(Chỉ Admin được thay đổi)*"
    )

  st.divider()
  st.header("➕ Thêm khoản chi")
  if bi_khoa:
    st.error("Tháng này đã bị khóa sổ.")
  else:
    ngay_dau_thang = hom_nay.replace(day=1)
    with st.form("form_chi_tieu", clear_on_submit=True):
      ngay = st.date_input(
          "Ngày chi",
          value=hom_nay,
          min_value=ngay_dau_thang,
          max_value=hom_nay,
      )
      danh_muc = st.selectbox(
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
      mo_ta = st.text_input("Mô tả (VD: Cà phê, Đổ xăng...)")
      so_tien = st.number_input("Số tiền (VNĐ)", min_value=1000, step=5000)
      st.caption(f"Người tạo: **{st.session_state.display_name}**")

      if st.form_submit_button("Lưu khoản chi", use_container_width=True):
        ws_chitieu.append_row([
            str(ngay),
            danh_muc,
            mo_ta if mo_ta else danh_muc,
            int(so_tien),
            st.session_state.display_name,
        ])
        st.rerun()

  if is_admin:
    st.divider()
    with st.expander("⚙️ Quản trị tài khoản (Admin)", expanded=False):
      member_keys = [u for u in users_db.keys() if u != "admin"]
      if not member_keys:
        st.write("Chưa có tài khoản thành viên nào khác.")
      else:
        for u_key in member_keys:
          u_info = users_db[u_key]
          u_col1, u_col2 = st.columns([2, 1])
          u_col1.write(f"• **{u_info['display_name']}** (`{u_key}`)")
          if u_col2.button("Xóa", key=f"del_user_{u_key}"):
            delete_user_from_sheet(u_key)
