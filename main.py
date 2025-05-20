from compute_initial_unique_crids import compute_unique_crids
from compute_initial_solution import compute_initial_solution
from add_initial_new_orders import add_new_orders


def main():
    print("Starting main process")
    print("1. Generate shell (No more use)")
    #compute_unique_crids()

    print("2. Construct initial solution (No more use)")
    #compute_initial_solution()

    print("3. Add new orders from initial snapshot (No more use)")
    add_initial_new_orders()



if __name__ == "__main__":
    main()