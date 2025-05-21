from A_A_compute_initial_unique_crids import compute_unique_crids
from A_B_compute_initial_solution import compute_initial_solution
from A_C_add_initial_new_orders import add_initial_new_orders
from A_D_update_solution_from_snapshots import update_solution_from_snapshots
from consolidate_heatmap_performance_sheets import consolidate_performance_sheets


def main():
    print("Starting main process")
    print("1. Generate shell (No more use)")
    #compute_unique_crids()

    print("2. Construct initial solution (No more use)")
    #compute_initial_solution()

    print("3. Add new orders from initial snapshot (No more use after Wednesday)")
    add_initial_new_orders()

    print("4. add new orders from future snapshots")

    #print("Consolidate heatmaps")
    #consolidate_performance_sheets()




if __name__ == "__main__":
    main()