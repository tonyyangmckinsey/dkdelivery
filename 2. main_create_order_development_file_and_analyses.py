from Produce_latest_Order_Development_Output import update_solution_from_snapshots
from Analyze_order_development import run_analysis
from Analyze_order_development_month import run_filtered_analysis

from Consolidate_heatmap_performance_sheets import consolidate_performance_sheets
# run this when we have the latest snapshots to create a new order development output, and to create the analyses

def main():
    print("1. Update the last Order Development output with the latest snapshots")
    update_solution_from_snapshots()
    print()
    print("2. Create the analyses")
    run_analysis()
    #run_filtered_analysis()
    print()
    print("3. Consolidate heat maps")
    consolidate_performance_sheets()
    print()

if __name__ == "__main__":
    main()