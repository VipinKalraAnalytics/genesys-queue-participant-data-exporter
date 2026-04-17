#!/usr/bin/env python3
"""
genesys_queue_participant_data_exporter.py

An interactive script for exporting participant data attributes from Genesys Cloud queue conversations to Excel.

This script authenticates with a Genesys Cloud OAuth client credentials app,
validates a queue, optionally enumerates available participant data attribute names,
runs an asynchronous analytics conversation details job filtered to the queue over
a given date range, extracts the selected participant data values, and writes
both deduplicated and raw results to separate sheets in an Excel workbook.

Typical permissions required:
- Analytics > Conversation Detail > View
- Analytics > Data Export > All
- Reporting > CustomParticipantAttributes > View

"""

import requests
import datetime
import time
import pandas as pd
import sys
import getpass
import os
from urllib.parse import quote_plus


class GenesysClient:
    """Client for interacting with the Genesys Cloud API."""

    def __init__(self, client_id: str, client_secret: str, region: str) -> None:
        self.client_id = client_id
        self.client_secret = client_secret
        self.region = region.strip()
        # Determine domain. If the region contains a dot, assume it's already a domain; otherwise
        # append .pure.cloud for AWS regions like 'usw2' or 'us-west-2'.
        if '.' in self.region:
            self.api_domain = self.region
        else:
            self.api_domain = f"{self.region}.pure.cloud"
        self.login_url = f"https://login.{self.api_domain}/oauth/token"
        self.api_base = f"https://api.{self.api_domain}"
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Genesys Exporter'})
        self.access_token: str | None = None

    def authenticate(self) -> None:
        """Authenticate using client credentials and store the bearer token."""
        data = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret
        }
        resp = self.session.post(self.login_url, data=data, timeout=30)
        resp.raise_for_status()
        token = resp.json().get('access_token')
        if not token:
            raise RuntimeError("Authentication response missing access_token")
        self.access_token = token
        self.session.headers.update({'Authorization': f'Bearer {token}'})

    def get_queue(self, queue_id: str) -> dict | None:
        """Retrieve queue details by ID. Returns None on failure."""
        url = f"{self.api_base}/api/v2/routing/queues/{quote_plus(queue_id)}"
        resp = self.session.get(url, timeout=30)
        if resp.status_code == 200:
            return resp.json()
        return None

    def submit_job(self, queue_id: str, start_iso: str, end_iso: str) -> str:
        """Submit an analytics conversation details job and return its ID."""
        url = f"{self.api_base}/api/v2/analytics/conversations/details/jobs"
        payload = {
            'interval': f'{start_iso}/{end_iso}',
            'filter': {
                'type': 'or',
                'predicates': [
                    {'dimension': 'queueId', 'value': queue_id}
                ]
            }
        }
        resp = self.session.post(url, json=payload, timeout=30)
        resp.raise_for_status()
        return resp.json().get('id')

    def get_job_status(self, job_id: str) -> str:
        """Check the status of an analytics conversation details job."""
        url = f"{self.api_base}/api/v2/analytics/conversations/details/jobs/{quote_plus(job_id)}"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json().get('state', '')

    def get_job_results(self, job_id: str) -> dict:
        """Retrieve the results of a completed analytics conversation details job."""
        url = f"{self.api_base}/api/v2/analytics/conversations/details/jobs/{quote_plus(job_id)}/results"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json()


def extract_attribute_names(results: dict) -> list[str]:
    """Extract a sorted list of participant attribute names from job results."""
    names: set[str] = set()
    for convo in results.get('conversations', []):
        for part in convo.get('participants', []):
            attrs = part.get('attributes') or {}
            names.update(attrs.keys())
    return sorted(names)


def flatten_results(results: dict, attr_name: str, queue_id: str) -> list[dict]:
    """Flatten job results to a list of records for the specified attribute and queue."""
    rows: list[dict] = []
    for convo in results.get('conversations', []):
        cid = convo.get('conversationId')
        start_time = convo.get('conversationStart')
        end_time = convo.get('conversationEnd')
        for part in convo.get('participants', []):
            attrs = part.get('attributes') or {}
            val = attrs.get(attr_name)
            if val is None:
                continue
            queue_segments = []
            for sess in part.get('sessions', []):
                for seg in sess.get('segments', []):
                    if seg.get('queueId') == queue_id:
                        queue_segments.append({'sessionId': sess.get('sessionId'), 'segment': seg})
            if queue_segments:
                for item in queue_segments:
                    seg = item['segment']
                    rows.append({
                        'conversationId': cid,
                        'conversationStart': start_time,
                        'conversationEnd': end_time,
                        'participantId': part.get('participantId'),
                        'sessionId': item['sessionId'],
                        'segmentStart': seg.get('segmentStart'),
                        'segmentEnd': seg.get('segmentEnd'),
                        'participantDataValue': val
                    })
            else:
                rows.append({
                    'conversationId': cid,
                    'conversationStart': start_time,
                    'conversationEnd': end_time,
                    'participantId': part.get('participantId'),
                    'sessionId': None,
                    'segmentStart': None,
                    'segmentEnd': None,
                    'participantDataValue': val
                })
    return rows


def deduplicate(rows: list[dict]) -> list[dict]:
    """Deduplicate rows by conversation ID and participant data value."""
    seen: set[tuple] = set()
    deduped: list[dict] = []
    for row in rows:
        key = (row['conversationId'], row['participantDataValue'])
        if key not in seen:
            seen.add(key)
            deduped.append(row)
    return deduped


def main() -> None:
    print("Genesys Queue Participant Data Exporter")
    print("=" * 50)
    print("This script exports participant data attributes from a Genesys Cloud queue into an Excel workbook.")
    print("You'll need an OAuth Client (Client ID + Secret) with the following typical permissions:")
    print(" - Analytics > Conversation Detail > View")
    print(" - Analytics > Data Export > All")
    print(" - Reporting > CustomParticipantAttributes > View")
    print()
    # Collect client credentials first
    client_id = input("Genesys Client ID: ").strip()
    client_secret = getpass.getpass("Genesys Client Secret: ").strip()
    # Region
    print("Enter your Genesys region (e.g., mypurecloud.com, usw2, us-west-2, mypurecloud.ie):")
    region = input("Genesys Region [mypurecloud.com]: ").strip() or "mypurecloud.com"
    client = GenesysClient(client_id, client_secret, region)
    print("Authenticating with Genesys Cloud...")
    try:
        client.authenticate()
    except Exception as exc:
        print(f"Authentication failed: {exc}")
        sys.exit(1)
    # Queue ID
    queue_id = input("Queue ID: ").strip()
    if not queue_id:
        print("Queue ID is required.")
        sys.exit(1)
    print("Validating queue...")
    queue = client.get_queue(queue_id)
    if not queue:
        print("Unable to retrieve queue details. Check the queue ID and your permissions.")
        sys.exit(1)
    print(f"Found queue: {queue.get('name')} (ID {queue.get('id')})")
    # Participant data name (attribute)
    attr_name = input("Participant Data Name (leave blank to discover): ").strip()
    # Number of days (default 1)
    days_input = input("Number of days [1]: ").strip()
    try:
        days = int(days_input) if days_input else 1
    except ValueError:
        days = 1
    # Output file name
    output_file = input("Output Excel file [genesys_queue_participant_data.xlsx]: ").strip() or "genesys_queue_participant_data.xlsx"
    # Determine date range: last N days
    end_dt = datetime.datetime.utcnow().replace(microsecond=0)
    start_dt = end_dt - datetime.timedelta(days=days)
    start_iso = start_dt.isoformat() + "Z"
    end_iso = end_dt.isoformat() + "Z"
    # Submit analytics job
    print(f"Submitting analytics job for {days} day(s)...")
    job_id = client.submit_job(queue_id, start_iso, end_iso)
    print(f"Job submitted with ID {job_id}. Waiting for completion (this may take several minutes)...")
    # Poll for completion
    while True:
        state = client.get_job_status(job_id)
        if state.lower() in ("complete", "completed"):
            break
        if state.lower() in ("failed", "error"):
            print(f"Analytics job ended with state: {state}")
            sys.exit(1)
        time.sleep(5)
    # Retrieve results
    print("Fetching job results...")
    results = client.get_job_results(job_id)
    # If attribute name not provided, discover names and prompt user
    if not attr_name:
        names = extract_attribute_names(results)
        if not names:
            print("No participant data attributes found in the conversation details.")
            sys.exit(0)
        print("Discovered participant data attributes:")
        for idx, name in enumerate(names, 1):
            print(f"{idx}. {name}")
        choice = input("Select attribute by number or enter name: ").strip()
        if choice.isdigit():
            index = int(choice)
            if 1 <= index <= len(names):
                attr_name = names[index - 1]
            else:
                print("Invalid selection.")
                sys.exit(1)
        else:
            attr_name = choice
        attr_name = attr_name.strip()
        if not attr_name:
            print("Attribute name is required.")
            sys.exit(1)
    # Flatten results for the chosen attribute
    rows = flatten_results(results, attr_name, queue_id)
    if not rows:
        print(f"No results found for attribute '{attr_name}'.")
        sys.exit(0)
    deduped = deduplicate(rows)
    # Create DataFrames
    df_raw = pd.DataFrame(rows)
    df_unique = pd.DataFrame(deduped)
    # Write to Excel
    print(f"Writing results to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_unique.to_excel(writer, sheet_name='Results_Unique', index=False)
        df_raw.to_excel(writer, sheet_name='Results_Detailed', index=False)
        meta = pd.DataFrame({
            'Metric': [
                'Selected Attribute',
                'Unique Rows',
                'Total Rows',
                'Start ISO',
                'End ISO',
                'Queue Name',
                'Queue ID'
            ],
            'Value': [
                attr_name,
                len(df_unique),
                len(df_raw),
                start_iso,
                end_iso,
                queue.get('name'),
                queue_id
            ]
        })
        meta.to_excel(writer, sheet_name='Metadata', index=False)
    print("Export complete.")
    print(f"Saved file: {output_file}")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("Exiting.")
