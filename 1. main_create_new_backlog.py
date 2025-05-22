from Backlog_Update__MergeAXBOwithcurrentbacklog import mergeaxbowithbacklog
from Backlog_Update_EnrichCurrentBacklogWithSteps import enrich_the_newest_merged_backlog_with_steps_from_the_latest_snapshot

# run this when axbo, Enterprisepublic, smb1, smb2, smb3 are updated

def main():
    print("1. Merging the axbo file with the latest backlog (Enterprisepublic, smb1, smb2, smb3)")
    mergeaxbowithbacklog()
    print("1.1 Latest backlog now merged with axbo")
    print()
    print("2. Enrich the latest merged backlog with the information from the latest DK backlog tracker (latest snapshot)")
    enrich_the_newest_merged_backlog_with_steps_from_the_latest_snapshot()
    print("2.1 Enriching done: Please update the DK backlog tracker with the Final_enriched_latest_backlog file in the folder 2. Output")
    print()

if __name__ == "__main__":
    main()