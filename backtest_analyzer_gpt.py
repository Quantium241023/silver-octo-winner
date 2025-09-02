import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import rc
import matplotlib
matplotlib.use('Agg')  # GUI 백엔드 문제 방지
import tkinter as tk
from tkinter import filedialog
import numpy as np
import os
import warnings
warnings.filterwarnings('ignore')  # 경고 메시지 숨기기
# 한글 폰트 지정 (예: 맑은 고딕, 나눔고딕 등)
rc('font', family='Malgun Gothic')   # 윈도우라면 'Malgun Gothic'
# 음수 기호 깨짐 방지
plt.rcParams['axes.unicode_minus'] = False

def main():
    print("=== 과최적화 분석 도구 (Net P&L % 기반) ===")
    
    # 1. 파일 선택
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="백테스트 엑셀 파일을 선택하세요",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        print("파일을 선택하지 않았습니다.")
        return
    
    print(f"선택한 파일: {file_path}")
    
    # 2. 엑셀 파일 읽기
    try:
        df_raw = pd.read_excel(file_path, sheet_name="List of trades")
    except Exception as e:
        print(f"엑셀 파일 읽기 오류: {e}")
        return
    
    print("\n=== 컬럼 목록 ===")
    print(df_raw.columns.tolist())
    
    # 3. 'Net P&L %' 및 '날짜' 컬럼 자동 감지
    # 수익률 컬럼 감지
    possible_pnl_pct_cols = [c for c in df_raw.columns if "p&l %" in c.lower() or "pnl %" in c.lower()]
    if not possible_pnl_pct_cols:
        pnl_pct_col = input("'Net P&L %'에 해당하는 컬럼명을 직접 입력하세요: ").strip()
        if pnl_pct_col not in df_raw.columns:
            print("수익률 컬럼을 찾을 수 없습니다."); return
    else:
        pnl_pct_col = possible_pnl_pct_cols[0]
    
    # 날짜 컬럼 감지 (Exit Time, Close Time 우선)
    possible_date_cols = [c for c in df_raw.columns if 'time' in c.lower() or 'date' in c.lower()]
    exit_time_cols = [c for c in possible_date_cols if 'exit' in c.lower() or 'close' in c.lower()]
    if exit_time_cols:
        date_col = exit_time_cols[0]
    elif possible_date_cols:
        date_col = possible_date_cols[0]
    else:
        print("날짜/시간에 해당하는 컬럼을 찾을 수 없습니다. (예: 'Exit Time', 'Close Time', 'Date')"); return

    print(f"\n분석에 사용할 수익률 컬럼: '{pnl_pct_col}'")
    print(f"분석에 사용할 날짜 컬럼: '{date_col}'")
    
    # 4. 거래 데이터 처리: 2행씩 짝지어서 하나의 거래로 처리
    trades_list = []
    data_rows = df_raw.reset_index(drop=True)
    
    for i in range(0, len(data_rows), 2):
        if i + 1 < len(data_rows):
            exit_row = data_rows.iloc[i + 1]
            
            trade_pnl_pct = exit_row[pnl_pct_col] if pd.notna(exit_row[pnl_pct_col]) else 0
            trade_date = exit_row[date_col]
            
            trades_list.append({
                'trade_id': i // 2 + 1,
                'pnl_pct': trade_pnl_pct,
                'date': trade_date
            })
        elif i < len(data_rows):
            print(f"⚠️ 짝이 없는 행 발견: {i+1}번째 행 (건너뜁니다)")

    if not trades_list:
        print("완결된 거래 데이터가 없습니다."); return
    
    # 5. 분석용 DataFrame 생성
    df = pd.DataFrame(trades_list)
    df['analysis_return'] = df['pnl_pct'] / 100.0
    df['date'] = pd.to_datetime(df['date'])
    
    print(f"\n추출된 완결 거래 수: {len(df)}")
    print(f"총 순수 수익률 합계: {df['analysis_return'].sum():.2%}")
    
    # 6. 과최적화 분석 시작
    print("\n" + "="*60)
    print("과최적화 분석 시작 (순수 수익률 기반)...")
    print("="*60)
    
    report = []
    n_trades = len(df)
    
    # 1) 완결된 거래 수
    report.append(f"완결된 거래 수 (진입+청산): {n_trades}")
    if n_trades < 30:
        report.append("⚠️ 완결된 거래 수가 너무 적습니다 (30 미만) → 과최적화 가능성")
    
    # 2) 소수 거래 집중도
    top_k = min(10, n_trades)
    total_return = df["analysis_return"].sum()
    top_returns = df["analysis_return"].nlargest(top_k).sum()
    
    if total_return > 0:
        ratio = top_returns / total_return
        report.append(f"상위 {top_k}개 거래의 수익률 비중: {ratio:.2%}")
        if ratio > 0.6:
            report.append("⚠️ 소수 거래에 성과 집중 → 과최적화 가능성")
    
    # 3) 월별 안정성 분석
    if n_trades > 0:
        monthly_returns = df.groupby(pd.Grouper(key='date', freq='M'))['analysis_return'].sum()
        
        report.append("\n--- 월별 수익률 분석 ---")
        for month, ret in monthly_returns.items():
            report.append(f"{month.strftime('%Y-%m')}: {ret:.2%}")
            
        negative_months = sum(1 for r in monthly_returns if r < 0)
        total_months = len(monthly_returns)
        if total_months > 0:
            report.append(f"\n총 {total_months}개월 중 {negative_months}개월이 음수 성과")
            if negative_months / total_months > 0.4: # 40% 이상이 음수면 경고
                report.append("⚠️ 음수 성과 월의 비율이 높음 → 불안정성")
        
        if total_months > 1:
            std_dev = np.std(monthly_returns)
            mean_return = np.mean(monthly_returns)
            cv = std_dev / abs(mean_return) if mean_return != 0 else float('inf')
            report.append(f"월별 수익률 변동계수: {cv:.2f}")
            if cv > 1.5:
                report.append("⚠️ 월별 성과 편차가 매우 큼 → 불안정성")
            
    # ... (이하 다른 분석들도 'analysis_return'을 사용하여 동일하게 수행) ...

    # 7. 결과 저장 및 그래프 생성
    out_dir = os.path.dirname(file_path)
    report_path = os.path.join(out_dir, "overfit_report_monthly.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("=== 과최적화 분석 리포트 (월별 분석 기반) ===\n\n")
        f.write("\n".join(report))
        
    try:
        fig, axes = plt.subplots(2, 1, figsize=(12, 8), gridspec_kw={'height_ratios': [2, 1]})
        
        # 누적 순수 수익률 곡선
        cumulative_returns = df['analysis_return'].cumsum() * 100
        axes[0].plot(range(1, len(df)+1), cumulative_returns, label="누적 순수 수익률", linewidth=2, color='purple')
        axes[0].set_title("누적 순수 수익률 곡선 (복리 효과 없음)")
        axes[0].set_xlabel("거래 번호")
        axes[0].set_ylabel("누적 수익률 (%)")
        axes[0].legend()
        axes[0].grid(True, alpha=0.3)
        
        # 월별 수익률
        if n_trades > 0 and not monthly_returns.empty:
            month_labels = monthly_returns.index.strftime('%Y-%m')
            colors = ['green' if r > 0 else 'red' for r in monthly_returns]
            axes[1].bar(month_labels, monthly_returns.values * 100, color=colors, alpha=0.7)
            axes[1].set_title("월별 수익률 분석")
            axes[1].set_xlabel("월(Month)")
            axes[1].set_ylabel("수익률 합계 (%)")
            plt.setp(axes[1].get_xticklabels(), rotation=45, ha="right")
            axes[1].grid(True, axis='y', alpha=0.3)
            axes[1].axhline(y=0, color='black', linestyle='-', alpha=0.5)
            
        plt.tight_layout()
        img_path = os.path.join(out_dir, "overfit_analysis_monthly.png")
        plt.savefig(img_path, dpi=150, bbox_inches='tight')
        plt.close()
        print(f"\n그래프 저장 완료: {img_path}")
        
    except Exception as e:
        print(f"\n그래프 생성 중 오류 발생: {e}")
        img_path = "그래프 생성 실패"

    # 8. 최종 결과 출력
    print("\n" + "="*60)
    print("과최적화 분석 결과 (월별 분석 기반)")
    print("="*60)
    for line in report:
        print(line)
    
    print(f"\n📊 리포트 저장: {report_path}")
    if img_path != "그래프 생성 실패":
        print(f"📈 그래프 저장: {img_path}")

if __name__ == "__main__":
    main()