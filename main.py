import requests
import shutil
import json
from requests_ntlm import HttpNtlmAuth  # Dùng cho xác thực Windows/NTLM
import os
from dotenv import load_dotenv
import pandas as pd
import openpyxl
import logging
import sys

# Cấu hình cơ bản cho logging
logging.basicConfig(
    level=logging.INFO,  # Mức độ log (INFO, DEBUG, WARNING, ERROR, CRITICAL)
    format='%(asctime)s - %(levelname)s - %(message)s',  # Định dạng log
    datefmt='%Y-%m-%d %H:%M:%S'  # Định dạng thời gian
)
logger = logging.getLogger(__name__)

def get_file_path(filename):
    if getattr(sys, 'frozen', False):  # Nếu đang chạy trong file .exe
        base_path = os.getcwd()
    else:  # Khi chạy Python script bình thường
        base_path = os.path.dirname(__file__)
    # try:
    #     base_path = os.path.dirname(os.path.abspath(__file__))
    # except NameError:  # Khi __file__ không tồn tại (vd: chạy từ .exe)
        # base_path = os.getcwd()  # Dùng thư mục hiện tại
    logger.info(f"File name is...{filename}")
    logger.info(f"Base path is...{base_path}")
    return os.path.join(base_path, filename)

# Sử dụng:
# config_path = get_file_path("PBI_CONFIG.xlsx")
env_path = get_file_path("PBI.env")

# Thông tin cấu hình
load_dotenv(env_path)
# Thông tin cấu hình
report_server_url = os.getenv("PBI_URL")
username = os.getenv("PBI_USER")
password = os.getenv("PBI_PWD")
config_env = os.getenv("PBI_CONFIG")
temp_env = os.getenv("PBI_CONFIG_TEMP")
config_security_sheet = os.getenv("PBI_SC_SHEET")
config_rls_sheet = os.getenv("PBI_RLS_SHEET")
policy_mapping = {
    "BROWSER": {"Name": "Browser", "Description": os.getenv("BROWSER")},
    "CONTENT_MANAGER": {"Name": "Content Manager", "Description": os.getenv("CONTENT_MANAGER")},
    "MY_REPORTS": {"Name": "My Reports", "Description": os.getenv("MY_REPORTS")},
    "PUBLISHER": {"Name": "Publisher", "Description": os.getenv("PUBLISHER")},
    "REPORT_BUILDER": {"Name": "Report Builder", "Description": os.getenv("REPORT_BUILDER")}
}


config_file = get_file_path(config_env)
temp_file = get_file_path(temp_env)

# dashboard_path = "/SANBOX/Dashboard/KHDN-sanbox/TEST_robotic/TEST_ROBOTIC_KHDN.DSH.002-FD.daily.report"  # Đường dẫn đến dashboard

def config_data(mode):
    report_dict = {}

    if mode == 'security':
        df = df_security
    if mode == 'rls':
        df = df_rls

    for _, row in df.iterrows():
        report_name = row["Report"]  # Lấy tên report từ cột report_name

        # Nếu report chưa có trong dictionary, khởi tạo
        if report_name not in report_dict:
            report_dict[report_name] = {
                "report_name": report_name,
                "path": row["Path"],
                "users": []
            }
        if mode == 'security':
        # Thêm thông tin người dùng vào danh sách users của report đó
            user_data = {
                "userid": row["User"],
                "permission": {
                    "BROWSER": row["Browser"],
                    "CONTENT_MANAGER": row["Content Manager"],
                    "MY_REPORTS": row["My Reports"],
                    "PUBLISHER": row["Publisher"],
                    "REPORT_BUILDER": row["Report Builder"],
                }
            }
        if mode == 'rls':
            permissions = row["Role"].split("#") if pd.notna(row["Role"]) else []
            user_data = {
                "userid": row["User"],
                "permission": permissions
            }
        
        report_dict[report_name]["users"].append(user_data)

    # Chuyển dictionary thành list
    report_list = list(report_dict.values())

    # In kết quả JSON
    print(json.dumps(report_list, indent=4))
    return report_list

# Hàm kiểm tra đăng nhập vào Power BI Report Server
def check_login():
    url = f"{report_server_url}/SystemResources?"  # Endpoint chung để kiểm tra xác thực
    
    # Gửi yêu cầu GET để kiểm tra đăng nhập
    response = requests.get(
        url,
        auth=HttpNtlmAuth(username, password),
        headers={"Content-Type": "application/json"}
    )
    
    if response.status_code == 200:
        logger.info(f"Kiểm tra đăng nhập: Thành công!")
        return True
    elif response.status_code == 401:
        raise Exception("Kiểm tra đăng nhập thất bại: Tài khoản hoặc mật khẩu không đúng (Unauthorized).")
    else:
        raise Exception(f"Kiểm tra đăng nhập thất bại: {response.status_code} - {response.text}")

# Hàm kiểm tra quyền truy cập
def check_access(dashboard_path):
    url = f"{report_server_url}/CatalogItems(Path='{dashboard_path}')"
    
    # Gửi yêu cầu thử nghiệm
    response = requests.get(
        url,
        auth=HttpNtlmAuth(username, password),
        headers={"Content-Type": "application/json"}
    )
    
    if response.status_code == 200:
        logger.info(f"Kiểm tra quyền truy cập: Thành công!")
        return True
    elif response.status_code == 401:
        raise Exception("Kiểm tra quyền truy cập thất bại: Tài khoản không được xác thực (Unauthorized).")
    elif response.status_code == 403:
        raise Exception("Kiểm tra quyền truy cập thất bại: Tài khoản không có quyền (Forbidden).")
    else:
        raise Exception(f"Kiểm tra quyền truy cập thất bại: {response.status_code} - {response.text}")
# Hàm lấy metadata của dashboard
def get_dashboard_metadata(dashboard_path):
    url = f"{report_server_url}/CatalogItems(Path='{dashboard_path}')"
    
    # Gửi yêu cầu GET đến API
    response = requests.get(
        url,
        auth=HttpNtlmAuth(username, password),
        headers={"Content-Type": "application/json"}
    )
    
    # Kiểm tra phản hồi
    if response.status_code == 200:
        metadata = response.json()
        return metadata
    else:
        raise Exception(f"Không thể lấy metadata: {response.status_code} - {response.text}")

# Hàm hiển thị metadata
def display_metadata(metadata):
    logger.info("Metadata của Dashboard:")
    logger.info(f"- ID: {metadata.get('Id')}")
    logger.info(f"- Tên: {metadata.get('Name')}")
    logger.info(f"- Đường dẫn: {metadata.get('Path')}")
    logger.info(f"- Loại: {metadata.get('Type')}")
    logger.info(f"- Ngày tạo: {metadata.get('CreatedDate')}")
    logger.info(f"- Ngày sửa đổi: {metadata.get('ModifiedDate')}")
    logger.info(f"- Kích thước: {metadata.get('Size')} bytes")
    logger.info(f"- Người tạo: {metadata.get('CreatedBy')}")
    logger.info(f"- Người sửa đổi cuối: {metadata.get('ModifiedBy')}")


# Hàm lấy ItemID của dashboard
def get_item_id(dashboard_path):
    url = f"{report_server_url}/CatalogItems(Path='{dashboard_path}')"
    response = requests.get(
        url,
        auth=HttpNtlmAuth(username, password),
        headers={"Content-Type": "application/json"}
    )
    
    if response.status_code == 200:
        item_data = response.json()
        return item_data["Id"]
    else:
        raise Exception(f"Không thể lấy ItemID: {response.status_code} - {response.text}")



# Hàm gán quyền
def assign_permissions(item_id, report):
    url = f"{report_server_url}/PowerBIReports({item_id})/Policies"
    # Lấy chính sách hiện tại
    response = requests.get(url, auth=HttpNtlmAuth(username, password))
    if response.status_code != 200:
        raise Exception(f"Không thể lấy chính sách hiện tại: {response.text}")
    
    policies_full = response.json()
    policies_full.pop('@odata.context', None)
    policies=response.json().get('Policies',[])
    # Payload định nghĩa chính sách phân quyền
    list_user = report['users']
    for user in list_user:
        group_name = "LPB\\"+user['userid']
        user_policies = user['permission']
        policy_list = [
        {"Name": policy_mapping[key]["Name"], "Description": policy_mapping[key]["Description"]}
        for key, value in user_policies.items() if value == 1
        ]

        has_group_name = any(policy['GroupUserName'] == group_name for policy in policies)

        if not policy_list and has_group_name:
            logger.info(f"Tiến hành xóa quyền Security cho user: {group_name}")
            policies = [policy for policy in policies if policy['GroupUserName'] != group_name]
            continue 

        if not policy_list:
            continue
        
        payload_new = {"GroupUserName": group_name, "Roles": policy_list}

        if has_group_name:
            for policy in policies:
                if policy['GroupUserName'] == group_name:
                    old_policy = policy['Roles']
                    if old_policy != policy_list:
                        logger.info(f"Tiến hành cập nhật Security cho user: {group_name}")
                        policy['Roles']=policy_list
        else:
            logger.info(f"Tiến hành thêm mới Security cho user: {group_name}")
            policies.append(payload_new)


    # print(policies)

    policies_full['Policies']=policies

    response = requests.request("PUT", url, headers={'Content-Type': 'application/json'}, data=json.dumps(policies_full, ensure_ascii=False).encode('utf8').decode(),auth=HttpNtlmAuth(username, password))
    
    if response.status_code == 200:
        # print(f"Đã gán quyền '{role_name}' cho '{group_name}' thành công!")
        logger.info(f"Đã gán quyền Security tự động thành công!")
        # print(f"Đã gán quyền tự động thành công!")
    else:
        raise Exception(f"Lỗi khi gán quyền: {response.status_code} - {response.text}")
    
# Hàm gán quyền
def assign_rls(item_id, report):
    url = f"{report_server_url}/PowerBIReports({item_id})/Policies"
    url_get_all_rls = f"{report_server_url}/PowerBIReports({item_id})/DataModelRoles"
    url_get_member_rls = f"{report_server_url}/PowerBIReports({item_id})/DataModelRoleAssignments"
    # Lấy chính sách hiện tại
    response_all_rls = requests.get(url_get_all_rls, auth=HttpNtlmAuth(username, password))
    response_member_rls = requests.get(url_get_member_rls, auth=HttpNtlmAuth(username, password))
    if response_all_rls.status_code != 200:
        raise Exception(f"Không thể lấy thông tin cấu hình rls của báo cáo: {response_all_rls.text}")
    if response_member_rls.status_code != 200:
        raise Exception(f"Không thể lấy rls người dùng của báo cáo: {response_member_rls.text}")
    
    rls_dict = report
    rls_policies=response_all_rls.json().get('value',[])
    rls_member_list = response_member_rls.json().get('value',[])
    role_mapping = {role['ModelRoleName']: role['ModelRoleId'] for role in rls_policies}

    payload = []

    for user in rls_dict["users"]:
        user_roles = [role_mapping[role] for role in user["permission"] if role in role_mapping]
        
        if user_roles:  # Chỉ thêm nếu có role hợp lệ
            payload.append({
                "GroupUserName": f"LPB\\{user['userid']}",
                "DataModelRoles": user_roles
            })
    # --- Bước 3: Cập nhật rls_member_list dựa trên payload ---
    # Chuyển rls_member_list thành dict để tra cứu nhanh theo GroupUserName
    rls_member_dict = {member["GroupUserName"]: member["DataModelRoles"] for member in rls_member_list}

    # Duyệt qua payload để cập nhật thông tin:
    for entry in payload:
        group = entry["GroupUserName"]
        new_roles = entry["DataModelRoles"]
        if group in rls_member_dict:
            # Nếu permission là "0" (tức danh sách rỗng) thì xóa user khỏi rls_member_dict
            if not new_roles:
                logger.info(f"Tiến hành xóa RLS cho user: {group}")
                del rls_member_dict[group]
            else:
                # Nếu có quyền thì cập nhật DataModelRoles với giá trị mới
                if sorted(rls_member_dict[group]) != sorted(new_roles):
                    logger.info(f"Tiến cập nhật RLS cho user: {group}")
                    rls_member_dict[group] = new_roles
        else:
            # Nếu GroupUserName chưa tồn tại và có quyền thì thêm mới vào
            if new_roles:
                logger.info(f"Tiến hành thêm mới RLS cho user: {group}")
                rls_member_dict[group] = new_roles

    # Chuyển kết quả từ dict về list
    updated_rls_member_list = [{"GroupUserName": k, "DataModelRoles": v} for k, v in rls_member_dict.items()]

    # In kết quả cuối cùng
    # print(json.dumps(updated_rls_member_list, indent=4))
    data=json.dumps(updated_rls_member_list, ensure_ascii=False).encode('utf8').decode()

    response = requests.request("PUT", url_get_member_rls, headers={'Content-Type': 'application/json'}, data=data,auth=HttpNtlmAuth(username, password))
    
    if response.status_code == 200:
        # print(f"Đã gán quyền '{role_name}' cho '{group_name}' thành công!")
        logger.info(f"Đã cập nhật RLS tự động thành công!")
        # print(f"Đã cập nhật RLS tự động thành công!")
    else:
        raise Exception(f"Lỗi khi cập nhật RLS: {response.status_code} - {response.text}")
    

def get_valid_mode():
    while True:
        print("Vui lòng chọn chế độ:")
        print("1. rls")
        print("2. security")
        
        choice = input("Nhập lựa chọn của bạn (1 hoặc 2): ")
        
        if choice == "1":
            return "rls"
        elif choice == "2":
            return "security"
        else:
            print("Lựa chọn không hợp lệ. Vui lòng chọn 1 hoặc 2.")



#LOAD_CONFIG_EXCEL
shutil.copy(config_file, temp_file)
with pd.ExcelFile(temp_file) as xls:
    rls_schema = {
    "User": str,        
    "Report": str,     
    "Path": str,        
    "Role": str       
    }
    xls = pd.ExcelFile(temp_file)
    print(xls.sheet_names)
    df_security = pd.read_excel(xls, sheet_name=config_security_sheet)
    df_rls = pd.read_excel(xls, sheet_name=config_rls_sheet, usecols=rls_schema.keys(),dtype=rls_schema)

check_login()

# Gọi hàm để lấy giá trị mode
mode = get_valid_mode()
print(f"Bạn đã chọn chế độ: {mode}")

# Sau đó sử dụng giá trị mode với hàm config_data
config_report_list = config_data(mode)

for report in config_report_list:
    dashboard_path = "/"+ report['path']
    try:
        check_access(dashboard_path)
        display_metadata(get_dashboard_metadata(dashboard_path))
        item_id = get_item_id(dashboard_path)
        logger.info(f"- Xử lý báo cáo với ID: {item_id}")
        if mode =='security':
            assign_permissions(item_id, report)
        if mode =='rls':
            assign_rls(item_id, report)
    except Exception as e:
        print(f"Đã xảy ra lỗi: {e}")

try:
    # Main logic của app
    logger.info(f"App is running...")
    
    # Nếu đang chạy ở chế độ .exe, giữ cửa sổ mở
    if getattr(sys, 'frozen', False):
        input("Press Enter to exit...")

except Exception as e:
    logger.info(f"Error: {e}")
    input("Press Enter to exit...")
