import pandas as pd
import subprocess
import re
import time
import requests
from packaging import version

# === Step 1: Load Excel Data ===
file_path = "2024_06_20_Software_Analysis_All.xlsx"  # Your uploaded file
df = pd.read_excel(file_path)

# Keep only relevant columns
df = df[['DisplayName', 'DisplayVersion']].dropna()
df.columns = ['DisplayName', 'DisplayVersion']

# Filter out undefined or empty values
df = df[
    (df['DisplayName'] != 'undefined') & 
    (df['DisplayName'].str.strip() != '') &
    (df['DisplayVersion'] != 'undefined') & 
    (df['DisplayVersion'].str.strip() != '')
]

# Drop duplicates
df = df.drop_duplicates(subset=['DisplayName'])

# Convert to list for processing
software_list = df.to_dict('records')

# === Step 2: Define Multiple Source Check Logic ===
def get_chocolatey_latest_version(software_name):
    """Get latest version from Chocolatey"""
    try:
        result = subprocess.run(
            ["choco", "info", software_name, "--limit-output"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            output = result.stdout.strip().split('\n')
            for line in output:
                if "|" in line:
                    parts = line.split("|")
                    if len(parts) >= 2:
                        return parts[1].strip()
        return None
    except Exception:
        return None

def get_winget_latest_version(software_name):
    """Get latest version from winget"""
    try:
        # Try direct search first
        result = subprocess.run(
            ["winget", "search", software_name, "-n", "3"],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            for line in lines:
                if software_name.lower() in line.lower():
                    # Parse winget output properly - version is usually in the 3rd column
                    parts = line.split()
                    if len(parts) >= 3:
                        # Check if the 3rd part looks like a version number
                        potential_version = parts[2]
                        if re.match(r'^\d+\.\d+', potential_version):
                            return potential_version
                    elif len(parts) >= 2:
                        # Fallback to 2nd part if it looks like a version
                        potential_version = parts[1]
                        if re.match(r'^\d+\.\d+', potential_version):
                            return potential_version
        
        # Try variations if direct search fails
        variations = [
            software_name.lower().replace(' ', ''),
            software_name.lower().replace(' ', '-'),
            software_name.lower().replace(' ', '.'),
            software_name.lower().split()[0]  # First word only
        ]
        
        for variation in variations:
            result = subprocess.run(
                ["winget", "search", variation, "-n", "3"],
                capture_output=True, text=True, timeout=5
            )
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                for line in lines:
                    if variation in line.lower():
                        # Parse winget output properly - version is usually in the 3rd column
                        parts = line.split()
                        if len(parts) >= 3:
                            # Check if the 3rd part looks like a version number
                            potential_version = parts[2]
                            if re.match(r'^\d+\.\d+', potential_version):
                                return potential_version
                        elif len(parts) >= 2:
                            # Fallback to 2nd part if it looks like a version
                            potential_version = parts[1]
                            if re.match(r'^\d+\.\d+', potential_version):
                                return potential_version
        return None
    except Exception:
        return None

def get_web_latest_version(software_name):
    """Get latest version from web sources"""
    try:
        # Common software websites to check
        search_urls = [
            f"https://chocolatey.org/packages?q={software_name.replace(' ', '+')}",
            f"https://winget.run/packages?q={software_name.replace(' ', '+')}",
            f"https://scoop.sh/#/apps?q={software_name.replace(' ', '+')}"
        ]
        
        for url in search_urls:
            try:
                response = requests.get(url, timeout=10, headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                })
                
                if response.status_code == 200:
                    content = response.text.lower()
                    # Look for version patterns in the page
                    version_patterns = [
                        r'(\d+\.\d+\.\d+)',  # x.x.x format
                        r'(\d+\.\d+)',       # x.x format
                        r'version[:\s]*(\d+\.\d+\.\d+)',
                        r'v(\d+\.\d+\.\d+)'
                    ]
                    
                    for pattern in version_patterns:
                        matches = re.findall(pattern, content)
                        if matches:
                            # Return the highest version found
                            versions = [version.parse(v) for v in matches if v]
                            if versions:
                                return str(max(versions))
                                
            except Exception:
                continue
        return None
    except Exception:
        return None

def get_latest_version_from_multiple_sources(software_name):
    """Get latest version from multiple sources in order of preference"""
    print(f"    Checking multiple sources for {software_name}...")
    
    # Source 1: Chocolatey
    latest = get_chocolatey_latest_version(software_name)
    if latest:
        print(f"      Chocolatey: {latest}")
        return latest
    
    # Source 2: Winget
    latest = get_winget_latest_version(software_name)
    if latest:
        print(f"      Winget: {latest}")
        return latest
    
    # Source 3: Web scraping
    latest = get_web_latest_version(software_name)
    if latest:
        print(f"      Web: {latest}")
        return latest
    
    print(f"      Not found in any source")
    return None

# === Step 3: Compare Versions ===
report = []
for i, entry in enumerate(software_list, 1):
    name = entry['DisplayName']
    installed = str(entry['DisplayVersion']).strip()

    print(f"[{i}/{len(software_list)}] Checking {name} (installed: {installed})...")

    latest = get_latest_version_from_multiple_sources(name)

    if latest:
        # Validate that latest looks like a version number
        if re.match(r'^\d+\.\d+', latest):
            try:
                # Handle different version formats
                installed_parsed = version.parse(installed)
                latest_parsed = version.parse(latest)
                
                if installed_parsed >= latest_parsed:
                    status = "Updated"
                else:
                    status = "Outdated"
            except:
                status = "Comparison Error"
        else:
            # If latest doesn't look like a version number, mark as unknown
            latest = "Invalid Version Format"
            status = "Unknown"
    else:
        latest = "Not Found"
        status = "Unknown"

    report.append({
        'Software Name': name,
        'Installed Version': installed,
        'Latest Version': latest,
        'Status': status
    })

    time.sleep(0.5)  # Brief pause between checks

# === Step 4: Save to Excel ===
report_df = pd.DataFrame(report)
output_path = "ACC_IMS_Third_Report.xlsx"
report_df.to_excel(output_path, index=False)

# Print summary
status_counts = report_df['Status'].value_counts()
print(f"\n📊 Summary:")
for status, count in status_counts.items():
    print(f"  {status}: {count}")

print(f"\n✅ Report saved to {output_path}")
