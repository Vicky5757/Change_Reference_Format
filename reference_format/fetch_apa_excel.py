import requests
import pandas as pd

def fetch_metadata(title):
    """
    Fetch metadata for a given title using the CrossRef API.
    """
    url = "https://api.crossref.org/works"
    params = {"query.title": title, "rows": 1}
    try:
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        if "message" in data and "items" in data["message"] and len(data["message"]["items"]) > 0:
            item = data["message"]["items"][0]

            # Process authors in APA format: Lastname, F. I.
            # Name formatting rules:
            # - For single author: "Lastname, F. I."
            # - For two authors: "Lastname1, F. I., & Lastname2, F. I."
            # - For three or more authors: "Lastname1, F. I., Lastname2, F. I., & Lastname3, F. I., et al."
            authors = []
            if "author" in item:
                for author in item["author"]:
                    given = author.get("given", "")  # First name(s)
                    family = author.get("family", "")  # Last name
                    if given and family:
                        # Convert given name(s) to initials
                        initials = " ".join([name[0] + "." for name in given.split()])
                        authors.append(f"{family}, {initials}")

            # Combine authors according to APA rules
            if len(authors) == 1:
                author_string = authors[0]
            elif len(authors) == 2:
                author_string = " & ".join(authors)
            elif len(authors) > 2:
                author_string = ", ".join(authors[:-1]) + ", & " + authors[-1]
            else:
                author_string = "N/A"

            metadata = {
                "Author": author_string,
                "Year": item.get("published-print", {}).get("date-parts", [[None]])[0][0] or "n.d.",
                "Title": title,
                "Journal": item.get("container-title", ["N/A"])[0] or "N/A",
                "Volume": item.get("volume", "N/A"),
                "Issue": item.get("issue", "N/A"),
                "Pages": item.get("page", "N/A"),
                "DOI": item.get("DOI", "N/A")
            }

            return metadata
        else:
            return {"Error": "No data found for this title."}
    except Exception as e:
        return {"Error": str(e)}

def format_apa(metadata):
    """
    Format metadata into APA 7th edition reference style.
    """
    if "Error" in metadata:
        return metadata["Error"]

    author = metadata["Author"]
    year = metadata["Year"]
    title = metadata["Title"]
    journal = metadata["Journal"]
    volume = metadata["Volume"]
    issue = f"({metadata['Issue']})" if metadata["Issue"] != "N/A" else ""
    pages = metadata["Pages"]
    doi = f"https://doi.org/{metadata['DOI']}" if metadata["DOI"] != "N/A" else "N/A"

    apa_reference = f"{author} ({year}). {title}. *{journal}*, {volume}{issue}, {pages}. {doi}"
    return apa_reference

def process_titles(input_csv, output_csv):
    """
    Process titles from an input CSV file and save APA references to an output CSV file.
    """
    try:
        df = pd.read_csv(input_csv)

        if "Title" not in df.columns:
            raise ValueError("Input CSV must contain a 'Title' column.")

        apa_references = []
        for title in df["Title"]:
            metadata = fetch_metadata(title)
            apa_reference = format_apa(metadata)
            apa_references.append(apa_reference)

        df["APA Reference"] = apa_references
        df.to_csv(output_csv, index=False)
        print(f"APA formatted references saved to {output_csv}")

    except FileNotFoundError as e:
        print(f"Error: {e}")
    except ValueError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    input_csv = "titles.csv"
    output_csv = "apa_references.csv"
    process_titles(input_csv, output_csv)