def sum_prefix_counts(filepath=r"C:\Users\Abc\Desktop\test.csv"):
    """
    Đọc file log, lấy 2 ký tự đầu mỗi dòng, tổng hợp số lượng của từng prefix 2 ký tự.
    Trả về dict {prefix: count}
    """
    prefix_counts = {}
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                line = line.rstrip("\n")
                if len(line) < 2:
                    continue
                prefix = line
                prefix_counts[prefix] = prefix_counts.get(prefix, 0) + 1
    except Exception as e:
        print(f"Lỗi khi đọc file {filepath}: {e}")
        return {}

    return prefix_counts

# Ví dụ sử dụng:
if __name__ == "__main__":
    result = sum_prefix_counts()
    for prefix, count in result.items():
        print(f"{prefix}: {count}")
    print(f"Tổng số mã prefix khác nhau: {len(result)}")
    print(f"Tổng số lượng mã prefix: {sum(result.values())}")