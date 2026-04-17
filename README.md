# Genesys Queue Participant Data Exporter

This repository contains a Python script that interactively exports custom participant data from a Genesys Cloud queue into an Excel workbook. It is intended for contact center administrators or analysts who need to retrieve values set through "Set Participant Data" actions in Genesys Architect flows.

## Features

- Authenticates to Genesys Cloud using your OAuth Client ID and Client Secret (Client Credentials grant).
- Prompts you for the Genesys region, queue ID, participant data attribute name, number of days to export, and output file name.
- Validates the queue ID and displays the queue name before continuing.
- Supports optional discovery of available participant data attribute names. If you leave the attribute name blank, the script will scan the selected queue/date range, list the attribute names found, and let you choose one.
- Submits an asynchronous Analytics Conversation Detail job filtered by queue and date range.
- Extracts the selected participant data attribute values from each conversation.
- Deduplicates the results by conversation ID + attribute value and writes both the unique and detailed results to separate sheets in the output workbook.
- Generates an Excel workbook with three sheets: **Results_Unique**, **Results_Detailed**, and **Metadata**.

## Prerequisites

- Python 3.7 or newer.
- A Genesys Cloud OAuth client (Client ID and Secret) with the following permissions assigned to the associated role(s):
  - **Analytics > Conversation Detail > View**
  - **Analytics > Data Export > All**
  - **Reporting > CustomParticipantAttributes > View**

  Without these, the script will not be able to submit analytics jobs or view participant attributes.
- Install the required Python packages:

```sh
pip install requests pandas openpyxl
```

## Usage

Run the script and follow the prompts:

```sh
python genesys_queue_participant_data_exporter.py
```

You will be asked for your Genesys Cloud client credentials, region, queue ID, participant data name, number of days to look back (default `1`) and the output filename. Leaving the attribute name blank triggers discovery of available participant data names and allows you to select one. If you already know the attribute name, enter it directly (the name is case ‑sensitive).

After validating the queue, the script submits an asynchronous conversation detail job for the specified queue and date range. When the job completes, it fetches the results, extracts the participant data attribute values, deduplicates them and writes an Excel file.

### Command ‑line arguments

Alternatively, you can supply parameters as command‑line arguments for automation or scheduling:

```sh
python genesys_queue_participant_data_exporter.py --client-id YOUR_CLIENT_ID --client-secret YOUR_CLIENT_SECRET --region mypurecloud.com --queue-id 12345678-1234-1234-1234-123456789012 --participant-data-name MyCustomAttribute --days 1 --output output.xlsx
```

Use the ISO region domain (e.g., `mypurecloud.com`, `usw2.pure.cloud`, `us-west-2` or `eu-west-1`) corresponding to your Genesys Cloud org. The script accepts common synonyms such as `usw2` and resolves them.

## Understanding the output

The script creates an Excel workbook with three sheets:

- **Results_Unique** – A deduplicated view containing one row per conversation ID and attribute value. Useful for a quick summary.
- **Results_Detailed** – A raw view of all matching queue segments and participants. A single conversation may produce multiple rows if it had multiple segments or participants.
- **Metadata** – Includes the job ID, queue ID, attribute name, date range and row counts for reference.

Keep in mind that Genesys Analytics jobs can take several minutes to complete, and data for very recent conversations may not be available immediately. Job results are typically aggregated hourly or daily by Genesys.

## Region examples

Genesys Cloud deployments reside in different AWS regions. Provide the region that matches your org:

- `mypurecloud.com` – North America (US East)
- `usw2.pure.cloud` or `us-west-2` – North America (US West, Oregon)
- `mypurecloud.ie` – Europe (Ireland)
- `mypurecloud.de` – Europe (Frankfurt)
- `mypurecloud.com.au` – Asia Pacific (Australia)

The script will build the appropriate API base URLs from the region you enter.

## Support

This script is provided as ‑is to help export participant data from Genesys Cloud queues. It uses only publicly documented Genesys Cloud APIs and does not require any specialized access beyond the permissions listed above.

Feel free to open an issue or pull request if you have suggestions or improvements.
