def monitor_relink(self, case_name):
    sanitized_case_name = self.sanitize_filename(case_name)
    output_file = self.logs_dir / f"log_{sanitized_case_name}.txt"
    try:
        start_time = datetime.now()
        end_time = start_time + timedelta(seconds=180)
        last_read_position = 0

        patterns_4_2s = ["08 02 01 0C", "08 02 01 41"]
        patterns_60s = ["08 02 01", "08 03 02", "08 01 03", "08 01 07", "08 01 0A"]

        match_4_2_sequence = []
        match_60s_sequence = []
        current_4_2_index = 0
        current_60_index = 0

        while datetime.now() <= end_time:
            if output_file.exists():
                with open(output_file, 'r') as file:
                    file.seek(last_read_position)
                    new_lines = file.readlines()
                    last_read_position = file.tell()

                    for line in new_lines:
                        line = line.strip()
                        if line:
                            if patterns_4_2s[current_4_2_index] in line:
                                match_4_2_sequence.append(datetime.now())
                                current_4_2_index += 1
                                if current_4_2_index == len(patterns_4_2s):
                                    current_4_2_index = 0

                            if patterns_60s[current_60_index] in line:
                                match_60s_sequence.append(datetime.now())
                                current_60_index += 1
                                if current_60_index == len(patterns_60s):
                                    current_60_index = 0

        valid_4_2s = len(match_4_2_sequence) >= len(patterns_4_2s) * 10
        valid_60s = len(match_60s_sequence) >= len(patterns_60s) * 2

        result_4_2s = "✅ Passed" if valid_4_2s else "❌ Failed"
        result_60s = "✅ Passed" if valid_60s else "❌ Failed"

        # Cập nhật UI từ luồng chính sau khi hoàn tất
        self.root.after(0, self.update_result, case_name, "✅ Completed", result_4_2s, result_60s)

    except Exception as e:
        print(f"[ERROR] Case {case_name}: Error during execution: {e}")
        self.root.after(0, self.update_result, case_name, "⚠️ Error", "N/A", "N/A")