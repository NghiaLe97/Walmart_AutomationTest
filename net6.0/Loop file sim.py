import os
import shutil


def copy_files_with_prefix(src_folder, i):
    # Lấy danh sách file trong thư mục nguồn
    files = [f for f in os.listdir(src_folder) if f.endswith('.sim')]

    for index in range(1, i + 1):
        for file_name in files:
            # Tạo đường dẫn đầy đủ của file gốc
            src_file = os.path.join(src_folder, file_name)

            # Tạo tên file mới với tiền tố (index)-
            new_file_name = f"({index})-{file_name}"
            dst_file = os.path.join(src_folder, new_file_name)

            # Copy file từ nguồn sang file mới
            shutil.copy2(src_file, dst_file)
            print(f"Copied {src_file} to {dst_file}")


# Ví dụ sử dụng:
src_folder = r"C:\Users\QuangMinh\PycharmProjects\performanceTest_androidTablet\Loop 1 Sim\New folder"  # Thay đường dẫn thư mục gốc ở đây
i = 20  # Nhập số lần muốn tạo bản sao
copy_files_with_prefix(src_folder, i)
