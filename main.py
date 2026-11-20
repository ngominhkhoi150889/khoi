from datetime import date
import hashlib
import io
import json
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Quản lý Chi tiêu", page_icon="💸", layout="wide"
)

FILE_DATA = "chi_tieu.csv"
FILE_BUDGET = "budget.json"
FILE_USERS = "users.json"


# --- HÀM BẢO MẬT & QUẢN LÝ TÀI KHOẢN ---
def hash_password(password):
  return hashlib.sha256(password.encode()).hexdigest()


def load_users():
  if not os.path.exists(FILE_USERS):
    default_users = {
        "admin": {
            "password": hash_password("SUPERACCOUNT@123"),
            "display_name": "Admin",
            "role": "admin",
        }
    }
    with open(FILE_USERS, "w", encoding="utf-8") as f:
      json.dump(default_users, f, ensure_ascii=False, indent=2)
    return default_users
  with open(FILE_USERS, "r", encoding="utf-8") as f:
    users = json.load(f)
    # Đảm bảo admin luôn có role là admin nếu dùng file cũ
    if "admin" in users and "role" not in users["admin"]:
      users["admin"]["role"] = "admin"
    return users


def save_users(users):
  with open(FILE_USERS, "w", encoding="utf-8") as f:
    json.dump(users, f, ensure_ascii=False, indent=2)


# --- HÀM XỬ LÝ DỮ LIỆU & NGÂN SÁCH ---
def load_data():
  if os.path.exists(FILE_DATA):
    df = pd.read_csv(FILE_DATA)
    df["Ngày"] = pd.to_datetime(df["Ngày"]).dt.date
    if "Người tạo" not in df.columns:
      df["Người tạo"] = "Admin"
    return df
  return pd.DataFrame(
      columns=["Ngày", "Danh mục", "Mô tả", "Số tiền", "Người tạo"]
  )


def save_data(df):
  df.to_csv(FILE_DATA, index=False)


def load_budget():
  if os.path.exists(FILE_BUDGET):
    with open(FILE_BUDGET, "r", encoding="utf-8") as f:
      return json.load(f).get("han_muc", 5000000)
  return 5000000


def save_budget(amount):
  with open(FILE_BUDGET, "w", encoding="utf-8") as f:
    json.dump({"han_muc": amount}, f)


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
        new_user = st.text_input("Tên đăng nhập mới (viết liền, không dấu)").strip()
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
            users_db[new_user] = {
                "password": hash_password(new_pass),
                "display_name": new_display,
                "role": "member",
            }
            save_users(users_db)
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

# Thanh tiêu đề & thông tin người dùng
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
        f"🔓 **Đang mở sổ tháng {thang_duoc_chon.split('-')[1]}/{thang_duoc_chon.split('-')[0]}.** Có thể nhập và xóa khoản nhầm."
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
      st.toast("Admin đã cập nhật hạn mức ngân sách!")
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
        new_row = pd.DataFrame([{
            "Ngày": ngay,
            "Danh mục": danh_muc,
            "Mô tả": mo_ta if mo_ta else danh_muc,
            "Số tiền": so_tien,
            "Người tạo": st.session_state.display_name,
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        save_data(df)
        st.rerun()

  # --- BẢNG ĐIỀU KHIỂN DÀNH RIÊNG CHO ADMIN ---
  if is_admin:
    st.divider()
    with st.expander("⚙️ Quản trị tài khoản (Admin)", expanded=False):
      st.caption("Danh sách người dùng hệ thống:")
      member_keys = [u for u in users_db.keys() if u != "admin"]
      if not member_keys:
        st.write("Chưa có tài khoản thành viên nào khác.")
      else:
        for u_key in member_keys:
          u_info = users_db[u_key]
          u_col1, u_col2 = st.columns([3, 1])
          u_col1.write(f"• **{u_info['display_name']}** (`{u_key}`)")
          if u_col2.button("Xóa", key=f"del_user_{u_key}"):
            del users_db[u_key]
            save_users(users_db)
            st.toast(f"Đã xóa tài khoản {u_key}")
            st.rerun()

# --- LỌC DỮ LIỆU THEO THÁNG & NGƯỜI TẠO ---
if not df.empty:
  df["Tháng_Năm"] = pd.to_datetime(df["Ngày"]).dt.strftime("%Y-%m")
  df_thang = df[df["Tháng_Năm"] == thang_duoc_chon].copy()
  if nguoi_tao_chon != "Tất cả":
    df_thang = df_thang[df_thang["Người tạo"] == nguoi_tao_chon]
else:
  df_thang = pd.DataFrame(
      columns=["Ngày", "Danh mục", "Mô tả", "Số tiền", "Người tạo"]
  )

st.divider()

# --- HIỂN THỊ CHỈ SỐ & CẢNH BÁO NGÂN SÁCH ---
tong_tien = df_thang["Số tiền"].sum() if not df_thang.empty else 0
con_lai = han_muc_thang - tong_tien
ty_le = tong_tien / han_muc_thang if han_muc_thang > 0 else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Tổng tiêu đã lọc", f"{tong_tien:,.0f} đ")
