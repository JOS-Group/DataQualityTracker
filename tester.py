from pathlib import Path
import re


# ============================================================
# FOLDER LOCATIONS
# ============================================================

FOLDER_1 = Path(
    r"C:\Users\Nick Holmes\OneDrive - Centers for Medicare Medicaid Services (CMS)\OIT-Workforce_Resilience - S\WR_LMS_Extracts"
)

FOLDER_2 = Path(
    r"C:\Users\Nick Holmes\OneDrive - Centers for Medicare Medicaid Services (CMS)\Peng, Thomas (CMS_CTR)'s files - source files\WR_LMS_Extracts"
)


# ============================================================
# NORMALIZE FILE NAME
# ============================================================

def normalize_filename(filename):
    """
    Normalizes a filename for comparison.

    Removes:
        - (1)
        - Workshop
        - Extra spaces

    Comparison is also case-insensitive.

    Example:
        "AI Training Workshop (1).xlsx"
        becomes
        "ai training.xlsx"
    """

    path = Path(filename)

    stem = path.stem
    extension = path.suffix.lower()

    # Remove "(1)" regardless of surrounding spaces
    stem = re.sub(r"\s*\(1\)\s*", " ", stem, flags=re.IGNORECASE)

    # Remove the word "Workshop"
    stem = re.sub(r"\bworkshop\b", " ", stem, flags=re.IGNORECASE)

    # Collapse multiple spaces into one
    stem = re.sub(r"\s+", " ", stem)

    # Remove leading/trailing spaces
    stem = stem.strip()

    # Case-insensitive comparison
    stem = stem.lower()

    return f"{stem}{extension}"


# ============================================================
# READ FOLDERS
# ============================================================

def get_files(folder):
    """
    Returns:
        normalized filename -> list of actual filenames

    A list is used in case multiple files normalize
    to the same name.
    """

    files = {}

    if not folder.exists():
        raise FileNotFoundError(f"Folder does not exist:\n{folder}")

    for item in folder.iterdir():
        if item.is_file():
            normalized = normalize_filename(item.name)

            if normalized not in files:
                files[normalized] = []

            files[normalized].append(item.name)

    return files


# ============================================================
# DISPLAY HELPERS
# ============================================================

def print_section(title):
    print("\n")
    print("=" * 80)
    print(title)
    print("=" * 80)


# ============================================================
# MAIN COMPARISON
# ============================================================

def main():

    print("Comparing LMS extract folders...")
    print()

    print("FOLDER 1:")
    print(FOLDER_1)

    print("\nFOLDER 2:")
    print(FOLDER_2)

    folder1_files = get_files(FOLDER_1)
    folder2_files = get_files(FOLDER_2)

    names1 = set(folder1_files.keys())
    names2 = set(folder2_files.keys())

    only_folder1 = sorted(names1 - names2)
    only_folder2 = sorted(names2 - names1)
    matching = sorted(names1 & names2)

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    print_section("SUMMARY")

    print(f"Files in Folder 1:            {sum(len(v) for v in folder1_files.values())}")
    print(f"Files in Folder 2:            {sum(len(v) for v in folder2_files.values())}")
    print(f"Matching normalized files:    {len(matching)}")
    print(f"Only in Folder 1:             {len(only_folder1)}")
    print(f"Only in Folder 2:             {len(only_folder2)}")

    # --------------------------------------------------------
    # ONLY IN FOLDER 1
    # --------------------------------------------------------

    print_section("FILES ONLY IN FOLDER 1")

    if not only_folder1:
        print("None")
    else:
        for normalized in only_folder1:
            for actual in folder1_files[normalized]:
                print(actual)

    # --------------------------------------------------------
    # ONLY IN FOLDER 2
    # --------------------------------------------------------

    print_section("FILES ONLY IN FOLDER 2")

    if not only_folder2:
        print("None")
    else:
        for normalized in only_folder2:
            for actual in folder2_files[normalized]:
                print(actual)

    # --------------------------------------------------------
    # MATCHES WHERE ACTUAL NAMES DIFFER
    # --------------------------------------------------------

    print_section("MATCHES WITH DIFFERENT ORIGINAL FILENAMES")

    different_name_matches = 0

    for normalized in matching:

        files1 = folder1_files[normalized]
        files2 = folder2_files[normalized]

        # If the actual filenames aren't identical,
        # show how they were matched.
        if set(files1) != set(files2):
            different_name_matches += 1

            print(f"\nNormalized as: {normalized}")

            print("  Folder 1:")
            for file in files1:
                print(f"    - {file}")

            print("  Folder 2:")
            for file in files2:
                print(f"    - {file}")

    if different_name_matches == 0:
        print("None")

    # --------------------------------------------------------
    # POSSIBLE DUPLICATES
    # --------------------------------------------------------

    print_section("POSSIBLE DUPLICATES AFTER NORMALIZATION")

    duplicate_found = False

    for folder_label, data in [
        ("Folder 1", folder1_files),
        ("Folder 2", folder2_files),
    ]:
        for normalized, actual_files in data.items():
            if len(actual_files) > 1:
                duplicate_found = True

                print(f"\n{folder_label} - {normalized}")

                for actual in actual_files:
                    print(f"    - {actual}")

    if not duplicate_found:
        print("None")

    print("\n")
    print("=" * 80)
    print("COMPARISON COMPLETE")
    print("=" * 80)


if __name__ == "__main__":
    main()