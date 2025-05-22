from A_A_compute_initial_unique_crids import compute_unique_crids
from A_B_compute_initial_solution import compute_initial_solution
from A_C_add_initial_new_orders import add_initial_new_orders
from Main_AA_Produce_Order_Development_Output import update_solution_from_snapshots
from Main_C_consolidate_heatmap_performance_sheets import consolidate_performance_sheets


def main():
    print("Starting main process")
    print("1. Generate shell (No more use)")
    #compute_unique_crids()

    print("2. Construct initial solution (No more use)")
    #compute_initial_solution()

    print("3. Add new orders from initial snapshot (No more use after Wednesday)(No more use)")
    #add_initial_new_orders()

    print("4. add new orders from future snapshots to produce Order_Devleopment_Output (Main algo)")
    #update_solution_from_snapshots()

    #print("Consolidate heatmaps")
    #consolidate_performance_sheets()


if __name__ == "__main__":
    main()