import cv2
import pandas as pd
import os

def extract_jump_the_difference():
    # --- 1. YOUR EXACT INPUTS ---
    xlsx_path = "Annom_WorkObject4.xlsx"
    output_dir = "workobject4_frames"  
    
    videos_info = [
        {
            "path": "video_data/20240521-160638190582.webm",
            "start": pd.to_datetime("2024-05-21 16:06:38.190")
        },
        {
            "path": "video_data/20240521-163107415027.webm",
            "start": pd.to_datetime("2024-05-21 16:31:07.415")
        },
        {
            "path": "video_data/20240522-101726215737.webm",
            "start": pd.to_datetime("2024-05-22 10:17:26.215")
        },
        {
            "path": "20240522-131943132767.webm",
            "start": pd.to_datetime("2024-05-22 13:19:43.132")
        },
        {
            "path": "video_data/reassembled_22.webm",
            "start": pd.to_datetime("2024-05-29 14:00:51.439")
        }
    ]

    print(f"Loading logs from {xlsx_path}...")
    df = pd.read_excel(xlsx_path)
    
    # Clean BOTH time columns
    df['Time'] = df['Time'].astype(str).str.replace(',', '.')
    df['Time'] = pd.to_datetime(df['Time'])
    
    df['Time1'] = df['Time1'].astype(str).str.replace(',', '.')
    df['Time1'] = pd.to_datetime(df['Time1'])
    
    # Sort by the LOGIC column (Time1) to guarantee we strictly move forward in the video
    df = df.sort_values(by='Time1') 
    
    os.makedirs(output_dir, exist_ok=True)
    total_extracted = 0

    # --- 2. PROCESS EACH VIDEO ---
    for i, vid in enumerate(videos_info):
        vid_path = vid['path']
        vid_start = vid['start']
        
        # Dynamic window: Assign logs to the correct video based on the LOGIC column (Time1)
        if i < len(videos_info) - 1:
            next_vid_start = videos_info[i+1]['start']
            vid_logs = df[(df['Time1'] >= vid_start) & (df['Time1'] < next_vid_start)]
        else:
            vid_logs = df[df['Time1'] >= vid_start]

        if vid_logs.empty:
            continue

        print(f"\n--- Opening Video {i+1}: {os.path.basename(vid_path)} ---")
        print(f"Assigned {len(vid_logs)} logs. Jumping differences using Time1...")
        
        cap = cv2.VideoCapture(vid_path)
        
        previous_log_time = vid_start 
        
        # --- 3. THE "JUMP THE DIFFERENCE" LOGIC ---
        for index, row in vid_logs.iterrows():
            # Split the tracks
            logic_time = row['Time1']  # Used for finding the frame
            name_time = row['Time']    # Used strictly for the file name
            
            # 1. Calculate the math using Time1
            jump_difference_ms = (logic_time - previous_log_time).total_seconds() * 1000.0
            target_ms = (logic_time - vid_start).total_seconds() * 1000.0
            
            if target_ms < 0:
                continue
            
            # 2. Fast-forward exactly that difference
            while True:
                current_ms = cap.get(cv2.CAP_PROP_POS_MSEC)
                
                # Did our fast-forward cross the gap?
                if current_ms >= target_ms:
                    ret, frame = cap.retrieve()
                    if ret:
                        # 3. Label the image using the original 'Time' column
                        label = "pore" if row['pore_diameter'] > 0 else "normal"
                        safe_time = name_time.strftime("%Y-%m-%d %H-%M-%S,%f")[:-3]
                        filename = f"{safe_time}_{label}.jpg"
                        
                        cv2.imwrite(os.path.join(output_dir, filename), frame)
                        total_extracted += 1
                        
                    # Update our anchor to this logic_time so we can calculate the NEXT difference
                    previous_log_time = logic_time 
                    break 
                
                # If we haven't crossed the gap yet, jump forward silently
                ret = cap.grab()
                if not ret:
                    break # Video ended

        cap.release()

    print(f"\nBoom! Finished. Saved {total_extracted} unique frames using Time1 logic and Time naming.")

extract_jump_the_difference()
