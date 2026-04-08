from ppt_optimizer import convert_emf_to_png_in_pptx, ProcessStatus

if __name__ == "__main__":
    INPUT = "./data/电力电子技术冲刺版课件7.pptx"
    OUTPUT = "./outdata/电力电子技术冲刺版课件7.pptx"
    MAX_WORKERS = 4
    
    print(f"开始处理，使用 {MAX_WORKERS} 个线程...")
    status, result = convert_emf_to_png_in_pptx(INPUT, OUTPUT, max_workers=MAX_WORKERS)
    
    if status == ProcessStatus.SUCCESS:
        print("\n处理完成")
        print(f"输入文件: {INPUT}")
        print(f"输出文件: {OUTPUT}")
        print(f"替换矢量图数量: {result}")
    elif status == ProcessStatus.FILE_NOT_FOUND:
        print(f"\n错误: {result}")
    elif status == ProcessStatus.ERROR_INTERRUPTED:
        print(f"\n处理中断: {result}")