import csv

def read_s2_blocks(file_path):
    """Read S2 CSV and extract J and K blocks."""
    print(f"Reading {file_path}")
    
    with open(file_path, encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        
        # Find J block (同行推荐人分析)
        in_j = False
        j_header = None
        j_data = []
        
        # Find K block (同行业绩分析 KEY ACCOUNT)
        in_k = False
        k_header = None
        k_data = []
        
        for row in reader:
            if not row:
                if in_j:
                    in_j = False
                if in_k:
                    in_k = False
                continue
            
            first = row[0].strip()
            
            # J block
            if first.startswith("J. 同行推荐人分析"):
                in_j = True
                continue
            if in_j and not j_header:
                j_header = [c.strip() for c in row]
                continue
            if in_j and j_header:
                if first.startswith("K."):
                    in_j = False
                    in_k = True
                    continue
                j_data.append(row)
                continue
            
            # K block
            if first.startswith("K. 同行业绩分析"):
                in_k = True
                continue
            if in_k and not k_header:
                k_header = [c.strip() for c in row]
                continue
            if in_k and k_header:
                if first.startswith("L-APE"):
                    break
                k_data.append(row)
                continue
    
    print("\n=== J Block (同行推荐人分析) ===")
    print(f"Header: {j_header}")
    print(f"Rows: {len(j_data)}")
    for row in j_data[:5]:
        print(f"  {row}")
    
    print("\n=== K Block (同行 KEY ACCOUNT) ===")
    print(f"Header: {k_header}")
    print(f"Rows: {len(k_data)}")
    for row in k_data[:10]:
        print(f"  {row}")

if __name__ == "__main__":
    read_s2_blocks("S2-业务端视角.csv")