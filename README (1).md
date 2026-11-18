# ストレート管理

Phần mềm quản lý nhập / xuất / tồn kho bằng QR (Excel VBA, Windows).

## Cách sao chép code
- **Trang có nút sao chép:** bật GitHub Pages (Settings → Pages → Branch: main → Save), mở `https://<tên-github>.github.io/<tên-repo>/`.
- **Hoặc trên GitHub:** bấm vào từng file `.txt` → nút **Copy raw file** (biểu tượng 2 tờ giấy, góc phải phía trên code).

## Thứ tự dán vào VBA

| # | File | Loại | Tên (Name) | Số dòng |
|---|---|---|---|---|
| 1 | `01_modUI.txt` | Module | `modUI` | 474 |
| 2 | `02_modData.txt` | Module | `modData` | 780 |
| 3 | `03_modApp.txt` | Module | `modApp` | 482 |
| 4 | `04_modSetup.txt` | Module | `modSetup` | 442 |
| 5 | `05_frmMainMenu.txt` | UserForm | `frmMainMenu` | 154 |
| 6 | `06_frmInbound.txt` | UserForm | `frmInbound` | 714 |
| 7 | `07_frmOutputSlip.txt` | UserForm | `frmOutputSlip` | 696 |
| 8 | `08_frmOutputProcess.txt` | UserForm | `frmOutputProcess` | 803 |
| 9 | `09_frmCancel.txt` | UserForm | `frmCancel` | 474 |
| 10 | `10_frmKumamoto.txt` | UserForm | `frmKumamoto` | 733 |
| 11 | `11_frmHistory.txt` | UserForm | `frmHistory` | 253 |
| 12 | `12_ThisWorkbook.txt` | ThisWorkbook | `ThisWorkbook` | 36 |

## Sau khi dán
1. Debug → Compile VBAProject
2. modSetup → đặt con trỏ trong `ストレート管理_初期設定` → F5 (tạo sheet)
3. Ctrl + S, đóng và mở lại file

UserForm để trống, không kéo thả control. Code đã bỏ sẵn dòng `Attribute VB_Name`.
