import requests
import json
from typing import Dict, List, Optional

# === CONFIGURATION ===
TOKEN = "7d373634e596a3694f57ceefb4011abb00280EE60CF600A645E315DD1BBF2BC293802A06"
BASE_URL = "https://hst-api.wialon.com/wialon/ajax.html"


# === WIALON API CALL ===
def wialon_api_call(svc: str, params=None, sid=None):
    """Generic Wialon API request helper."""
    payload = {"svc": svc}
    if params:
        payload["params"] = json.dumps(params)
    if sid:
        payload["sid"] = sid
    response = requests.get(BASE_URL, params=payload)
    return response.json()


# === LOGIN ===
def login_with_token(token: str) -> str:
    """Authenticate using token and return SID."""
    data = wialon_api_call("token/login", {"token": token})
    sid = data.get("eid") or data.get("sid")
    if not sid:
        raise Exception(f"Failed to log in: {data}")
    print("✅ Logged in. SID:", sid)
    return sid


# === SEARCH RESOURCES (for report templates) ===
def search_resources_standalone(sid: str) -> Optional[List[Dict]]:
    """Fetch all resources that may contain report templates."""
    params = {
        "spec": {
            "itemsType": "avl_resource",
            "propName": "reporttemplates",
            "propValueMask": "*",
            "sortType": "reporttemplates"
        },
        "force": 1,
        "flags": 8193,  # basic info + report templates
        "from": 0,
        "to": 0
    }
    try:
        response = wialon_api_call("core/search_items", params, sid)
        items = response.get("items")
        if not items:
            print("⚠️ No resources found.")
            return None
        return items
    except Exception as e:
        print(f"❌ Error searching resources: {e}")
        return None


# === GET ALL TEMPLATES ===
def get_all_templates_standalone(sid: str) -> List[Dict]:
    """List all report templates from available resources."""
    resources = search_resources_standalone(sid)
    if not resources:
        return []

    all_templates = []
    for resource in resources:
        resource_name = resource.get("nm", "Unknown Resource")
        resource_id = resource.get("id")
        if "rep" in resource:
            for template_id, template_data in resource["rep"].items():
                template_info = {
                    "resource_name": resource_name,
                    "resource_id": resource_id,
                    "template_id": template_id,
                    "template_name": template_data.get("n", "Unnamed Template"),
                    "template_type": "Group" if template_data.get("ct") == "avl_unit_group" else "Single Unit"
                }
                all_templates.append(template_info)
    return all_templates


# === SAVE TEMPLATES TO FILE ===
def save_templates_to_file_standalone(sid: str, filename: str = "wialon_templates.json"):
    """Save all templates to a JSON file."""
    templates = get_all_templates_standalone(sid)
    if templates:
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(templates, f, indent=2, ensure_ascii=False)
        print(f"\n💾 Templates saved to {filename}")


# === LIST UNITS ===
def list_units(sid: str) -> List[Dict]:
    """List all units (trucks, gensets, etc.)."""
    params = {
        "spec": {
            "itemsType": "avl_unit",
            "propName": "sys_name",
            "propValueMask": "*",
            "sortType": "sys_name"
        },
        "force": 1,
        "flags": 1,
        "from": 0,
        "to": 0
    }
    res = wialon_api_call("core/search_items", params, sid)
    units = res.get("items", [])
    print("\n🚛 Available Units:")
    for u in units:
        print(f"Unit: {u['nm']}, ID: {u['id']}")
    return [{"name": u["nm"], "id": u["id"]} for u in units]


# === LIST UNIT GROUPS ===
def list_unit_groups(sid: str) -> List[Dict]:
    """List all unit groups (e.g., TRANSIT ALL TRUCKS, COPPER TRUCKS)."""
    params = {
        "spec": {
            "itemsType": "avl_unit_group",
            "propName": "sys_name",
            "propValueMask": "*",
            "sortType": "sys_name"
        },
        "force": 1,
        "flags": 1,
        "from": 0,
        "to": 0
    }
    res = wialon_api_call("core/search_items", params, sid)
    groups = res.get("items", [])
    print("\n📦 Available Unit Groups:")
    for g in groups:
        print(f"Group: {g['nm']}, ID: {g['id']}")
    return [{"name": g["nm"], "id": g["id"]} for g in groups]


# === MAIN FUNCTION ===
def main():
    sid = login_with_token(TOKEN)

    # Get data
    units = list_units(sid)
    groups = list_unit_groups(sid)
    templates = get_all_templates_standalone(sid)

    # Save separate files
    with open("wialon_units.json", "w", encoding="utf-8") as f:
        json.dump(units, f, indent=2, ensure_ascii=False)
    print("💾 Units saved to wialon_units.json")

    with open("wialon_groups.json", "w", encoding="utf-8") as f:
        json.dump(groups, f, indent=2, ensure_ascii=False)
    print("💾 Groups saved to wialon_groups.json")

    save_templates_to_file_standalone(sid)

    # Combined overview file
    ids_data = {"units": units, "groups": groups, "templates": templates}
    with open("wialon_ids_overview.json", "w", encoding="utf-8") as f:
        json.dump(ids_data, f, indent=2, ensure_ascii=False)
    print("\n💾 All IDs (units, groups, templates) saved to wialon_ids_overview.json")


if __name__ == "__main__":
    main()
