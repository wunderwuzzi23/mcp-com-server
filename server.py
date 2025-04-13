import win32com.client
import uuid
import sys
import logging
from typing import Dict, Any, List, Optional, Union
from mcp.server.fastmcp import FastMCP

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("mcp-com-server")

# HRESULT codes (commonly used in COM)
S_OK = 0x00000000          # Operation successful
S_FALSE = 0x00000001         # Successful but returns a false condition
E_FAIL = 0x80004005          # Unspecified failure
E_NOINTERFACE = 0x80004002     # Interface not supported
E_INVALIDARG = 0x80070057      # One or more arguments are invalid
E_ACCESSDENIED = 0x80070005      # Access denied
DISP_E_MEMBERNOTFOUND = 0x80020003 # Member not found

def hr_to_string(hr: int) -> str:
    """Convert HRESULT code to a readable string."""
    hresult_map = {
        S_OK: "S_OK (0x00000000) - Operation successful",
        S_FALSE: "S_FALSE (0x00000001) - Successful but false condition",
        E_FAIL: "E_FAIL (0x80004005) - Unspecified failure",
        E_NOINTERFACE: "E_NOINTERFACE (0x80004002) - Interface not supported",
        E_INVALIDARG: "E_INVALIDARG (0x80070057) - One or more arguments are invalid",
        E_ACCESSDENIED: "E_ACCESSDENIED (0x80070005) - Access denied",
        DISP_E_MEMBERNOTFOUND: "DISP_E_MEMBERNOTFOUND  (0x80020003) - Member not found"
    }
    return hresult_map.get(hr, f"Unknown HRESULT: {hex(hr)}")

# Initialize the MCP server
mcp = FastMCP("mcp-com-server", version="1.0", description="MCP COM Server for Agents")

# Object registry to track all created COM objects - storing both objects and their ProgIDs
object_registry: Dict[str, Dict[str, Any]] = {}

# ALLOWLIST for COM objects (empty means all allowed)
COM_ALLOWLIST = []

def is_com_object_allowed(identifier: str) -> bool:
    """Check if a COM object is allowed to be created based on security ALLOWLIST."""
    # If ALLOWLIST is empty, allow all objects (for development)
    if not COM_ALLOWLIST:
        return True
    return identifier in COM_ALLOWLIST

def get_type_info(obj) -> Dict[str, Any]:
    """
    Get extended type information about a COM object.
    This function attempts to discover available methods and properties of a COM object
    through introspection, providing additional metadata where possible.
    
    Args:
        obj: COM object to inspect
        
    Returns:
        Dictionary containing categorized methods and properties with metadata
    """
    type_info = {
        "methods": [],
        "properties": [],
        "events": [],
        "errors_encountered": []
    }

    # Include both public and private members (COM objects often expose important
    # functionality through "_" prefixed members)
    for name in dir(obj):
        try:
            # For COM objects, accessing some attributes might raise exceptions
            # even during simple inspection
            attr = getattr(obj, name)
            
            # Try to determine if it's a method or property
            if callable(attr):
                type_info["methods"].append({
                    "name": name,
                    "is_private": name.startswith("_"),
                    # Try to get signature if available
                    "signature": str(getattr(attr, "__signature__", "Unknown"))
                })
            else:
                # It's a property
                if name.startswith("on") and name[2:3].isupper():
                    # Heuristic for identifying potential event handlers
                    type_info["events"].append(name)
                else:
                    property_type = "Unknown"
                    try:
                        property_type = type(attr).__name__
                    except:
                        pass
                        
                    type_info["properties"].append({
                        "name": name,
                        "type": property_type,
                        "is_private": name.startswith("_"),
                        "is_com_object": hasattr(attr, "_oleobj_") if attr is not None else False
                    })
        except Exception as e:
            # Log specific errors without breaking the entire inspection
            type_info["errors_encountered"].append({
                "member": name,
                "error": str(e)
            })
            
    return type_info

@mcp.tool("CreateObject")
def co_create_instance(identifier: str) -> Dict[str, Any]:
    """
    This tool can be used to leverage Windows COM Automation to use Office applications, spreadsheets, 
    emails, and anything else that is based on COM. It uses the win32com.client (python wrapper) module 
    to create COM objects via a CLSID or ProgID (CoCreateInstance). 
    The tool returns a unique runtime_id for the created object, which can be used to interact with it. 
    For properties that are objects, a new runtime_id object is returned that you can use to interact. 
    The same happens if an InvokeMethod returns a COM object, you can use the new runtime_id to interact with it.
    """
    if not is_com_object_allowed(identifier):
        result = E_ACCESSDENIED
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: {identifier} is not on the ALLOWLIST",
            "runtime_id": None
        }

    try:
        # Create the COM object using Dispatch, which internally calls CoCreateInstance
        com_object = win32com.client.Dispatch(identifier)

        # Generate a unique ID for this object instance
        runtime_id = str(uuid.uuid4())

        # Determine if the identifier is a ProgID or CLSID
        # If it starts with { and ends with }, it's likely a CLSID
        is_clsid = identifier.startswith('{') and identifier.endswith('}')

        clsid = None
        prog_id = None

        if is_clsid:
            clsid = identifier
            # Try to get ProgID from CLSID (may not always work)
            try:
                prog_id = win32com.client.ProgIDFromCLSID(identifier)
            except:
                prog_id = "Unknown"
        else:
            prog_id = identifier
            # Try to get CLSID from ProgID
            try:
                clsid = win32com.client.GetCLSIDFromProgID(identifier)
            except:
                clsid = "Unknown"

        # Store the COM object and both identifiers in our registry
        object_registry[runtime_id] = {
            "object": com_object,
            "prog_id": prog_id,
            "clsid": clsid
        }

        result = S_OK
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Successfully created COM object: {identifier}",
            "runtime_id": runtime_id
        }
    except Exception as e:
        result = E_FAIL
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Failed to create COM object: {str(e)}",
            "runtime_id": None
        }

@mcp.tool("QueryInterface")
def query_interface(runtime_id: str, iid: str) -> Dict[str, Any]:
    """
    Retrieves a pointer to a specified interface of a COM object and returns a new runtime ID.
    """
    if runtime_id not in object_registry:
        result = E_INVALIDARG
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Invalid object ID: {runtime_id}",
            "runtime_id": None
        }

    com_object = object_registry[runtime_id]["object"]

    try:
        interface = com_object.QueryInterface(iid)
        new_runtime_id = str(uuid.uuid4())
        # Try to get ProgID/CLSID for the new interface (may not always be possible)
        prog_id = "Unknown"
        clsid = "Unknown"
        try:
            clsid_bytes = interface._oleobj_.GetCLSID()
            clsid = str(uuid.UUID(bytes_le=clsid_bytes))
            prog_id = win32com.client.ProgIDFromCLSID(clsid)
        except:
            pass

        object_registry[new_runtime_id] = {
            "object": interface,
            "prog_id": prog_id,
            "clsid": clsid
        }

        result = S_OK
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Successfully queried interface: {iid}. New runtime ID: {new_runtime_id}",
            "runtime_id": new_runtime_id
        }
    except Exception as e:
        result = E_NOINTERFACE
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Interface not supported: {str(e)}",
            "runtime_id": None
        }

@mcp.tool("GetTypeInformation")
def get_type_information(runtime_id: str) -> Dict[str, Any]:
    """
    Provides extended type information about a COM object.
    Windows COM API equivalent: ITypeInfo::GetTypeAttr, ITypeInfo::GetDocumentation
    """
    if runtime_id not in object_registry:
        result = E_INVALIDARG
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Invalid object ID: {runtime_id}",
            "type_info": None
        }

    com_object = object_registry[runtime_id]["object"]

    try:
        # Get extended type information
        type_info = get_type_info(com_object)

        result = S_OK
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Successfully retrieved type information",
            "type_info": type_info
        }
    except Exception as e:
        result = E_FAIL
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Failed to get type information: {str(e)}",
            "type_info": None
        }

@mcp.tool("InvokeMethod")
def invoke_method(runtime_id: str, method_name: str, args=None) -> Dict[str, Any]:
    """
    Invokes a method on a previously registered COM object (runtime_id)
    This function wraps the IDispatch::Invoke method.
    Calling Conventions:
    - Use Navigate2, instead of Navigate for IE.
    - Instead of empty args (None), try using an empty list [] first.
    """
    if runtime_id not in object_registry:
        result = E_INVALIDARG
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Invalid object ID: {runtime_id}",
            "return_value": None
        }

    com_object = object_registry[runtime_id]["object"]

    try:
        # Check if method exists
        if not hasattr(com_object, method_name):
            result = DISP_E_MEMBERNOTFOUND
            return {
                "result": result,
                "message": f"{hr_to_string(result)}: Method not found: {method_name}",
                "return_value": None
            }

        # Fix for methods that require arguments but None is provided
        if args is None:
            args = []

        # Invoke the method
        try:
            method = getattr(com_object, method_name)
            return_value = method(*args)
            result = S_OK
            
            # Check if return value is a COM object by type
            is_com_object = False
            
            # Check for common win32com class types
            if (hasattr(return_value, "_oleobj_") or
                type(return_value).__name__ == 'CDispatch' or 
                'win32com' in str(type(return_value).__module__) or
                '<COMObject' in str(return_value)):
                is_com_object = True
            
            # Register if it's a COM object
            if is_com_object:
                new_runtime_id = str(uuid.uuid4())
                clsid = "Unknown"
                prog_id = "Unknown"
                
                # Try to get CLSID if possible (for informational purposes)
                if hasattr(return_value, "_oleobj_"):
                    try:
                        clsid_bytes = return_value._oleobj_.GetCLSID()
                        clsid = str(uuid.UUID(bytes_le=clsid_bytes))
                        
                        # Try to get ProgID
                        try:
                            prog_id = win32com.client.ProgIDFromCLSID(clsid)
                        except Exception as e:
                            logger.debug(f"Failed to get ProgID: {e}")
                    except Exception as e:
                        logger.debug(f"Failed to get CLSID: {e}")
                
                # Register the COM object
                object_registry[new_runtime_id] = {
                    "object": return_value,
                    "prog_id": prog_id,
                    "clsid": clsid
                }
                
                return {
                    "result": result,
                    "message": f"{hr_to_string(result)}: Successfully invoked method: {method_name} and registered return value as COM object. Reference it with runtime_id: {new_runtime_id}",
                    "return_value": new_runtime_id
                }
            else:
                # Not a COM object, return the value directly
                return {
                    "result": result,
                    "message": f"{hr_to_string(result)}: Successfully invoked method: {method_name}",
                    "return_value": return_value
                }
                
        except Exception as e:
            result = E_FAIL
            return {
                "result": result,
                "message": f"{hr_to_string(result)}: Failed to invoke method '{method_name}': {str(e)}",
                "return_value": None
            }
    except Exception as e:
        result = E_FAIL
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: An unexpected error occurred while invoking method '{method_name}': {str(e)}",
            "return_value": None
        }

@mcp.tool("GetProperty")
def get_property(runtime_id: str, property_name: str) -> Dict[str, Any]:
    """
    Gets a property value from a COM object. It attempts to register the returned object as a
    COM object by checking its type. If this fails, it's treated as a regular property value.
    Windows COM API equivalent: IDispatch::Invoke (DISPATCH_PROPERTYGET)
    """
    if runtime_id not in object_registry:
        result = E_INVALIDARG
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Invalid object ID: {runtime_id}",
            "value": None
        }

    com_object = object_registry[runtime_id]["object"]

    try:
        # Check if property exists
        if not hasattr(com_object, property_name):
            result = DISP_E_MEMBERNOTFOUND
            return {
                "result": result,
                "message": f"{hr_to_string(result)}: Property not found: {property_name}",
                "value": None
            }

        try:
            # Get the property value
            value = getattr(com_object, property_name)
            result = S_OK
            
            # Check if it's a COM object by type
            # Import the specific types at the top of your file:
            # from win32com.client import CDispatch, CoClassBaseClass, DispatchBaseClass
            is_com_object = False
            
            # Check for common win32com class types
            if (type(value).__name__ == 'CDispatch' or 
                isinstance(value, (win32com.client.CDispatch, 
                                  win32com.client.CoClassBaseClass,
                                  win32com.client.DispatchBaseClass)) or
                'win32com.gen_py' in str(type(value).__module__)):
                is_com_object = True
            
            # Register if it's a COM object
            if is_com_object:
                new_runtime_id = str(uuid.uuid4())
                clsid = "Unknown"
                prog_id = "Unknown"
                
                # Try to get CLSID if possible (for informational purposes)
                if hasattr(value, "_oleobj_"):
                    try:
                        clsid_bytes = value._oleobj_.GetCLSID()
                        clsid = str(uuid.UUID(bytes_le=clsid_bytes))
                        
                        # Try to get ProgID
                        try:
                            prog_id = win32com.client.ProgIDFromCLSID(clsid)
                        except Exception as e:
                            logger.debug(f"Failed to get ProgID: {e}")
                    except Exception as e:
                        logger.debug(f"Failed to get CLSID: {e}")
                
                # Register the COM object
                object_registry[new_runtime_id] = {
                    "object": value,
                    "prog_id": prog_id,
                    "clsid": clsid
                }
                
                return {
                    "result": result,
                    "message": f"{hr_to_string(result)}: Successfully got property: {property_name} and registered COM object. Reference it with runtime_id: {new_runtime_id}",
                    "value": new_runtime_id
                }
            else:
                # Not a COM object, return the value directly
                return {
                    "result": result,
                    "message": f"{hr_to_string(result)}: Successfully got property: {property_name}",
                    "value": value
                }
            
        except Exception as e:
            result = E_FAIL
            return {
                "result": result,
                "message": f"{hr_to_string(result)}: Failed to get property '{property_name}': {str(e)}",
                "value": None
            }
    except Exception as e:
        result = E_FAIL
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: An unexpected error occurred while getting property '{property_name}': {str(e)}",
            "value": None
        }

@mcp.tool("SetProperty")
def set_property(runtime_id: str, property_name: str, value: Any) -> Dict[str, Any]:
    """
    Sets a property value on a COM object.
    Windows COM API equivalent: IDispatch::Invoke (DISPATCH_PROPERTYPUT)
    """
    if runtime_id not in object_registry:
        result = E_INVALIDARG
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Invalid object ID: {runtime_id}"
        }

    com_object = object_registry[runtime_id]["object"]

    try:
        # Set the property
        if not hasattr(com_object, property_name):
            result = DISP_E_MEMBERNOTFOUND
            return {
                "result": result,
                "message": f"{hr_to_string(result)}: Property not found: {property_name}"
            }

        setattr(com_object, property_name, value)

        result = S_OK
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Successfully set property: {property_name}"
        }
    except Exception as e:
        result = E_FAIL
        return {
            "result": result,
            "message": f"{hr_to_string(result)}: Failed to set property: {str(e)}"
        }

@mcp.tool("DisposeObject")
def dispose_object(runtime_id_or_ids: Union[str, List[str]]) -> Dict[str, Any]:
    """
    Disposes one or more COM objects and removes them from the registry.
    Typically only dispose of objects if the user wants to close a specific application or interaction,
    otherwise keep things around. Windows COM API equivalent: IUnknown::Release
    
    Args:
        runtime_id_or_ids: Either a single runtime_id string or a list of runtime_id strings
        
    Returns:
        A dictionary containing result code and status message
    """
    # Handle single ID case
    if isinstance(runtime_id_or_ids, str):
        runtime_id_or_ids = [runtime_id_or_ids]
    
    # Track results for each object
    results = []
    overall_result = S_OK
    
    # Process each object ID
    for runtime_id in runtime_id_or_ids:
        if runtime_id not in object_registry:
            result = E_INVALIDARG
            results.append({
                "id": runtime_id,
                "result": result,
                "message": f"{hr_to_string(result)}: Invalid object ID: {runtime_id}"
            })
            overall_result = E_FAIL
            continue
            
        try:
            # Remove the object from the registry
            # Python's garbage collection will handle COM reference counting
            del object_registry[runtime_id]
            result = S_OK
            results.append({
                "id": runtime_id,
                "result": result,
                "message": f"{hr_to_string(result)}: Successfully disposed object"
            })
        except Exception as e:
            result = E_FAIL
            results.append({
                "id": runtime_id,
                "result": result,
                "message": f"{hr_to_string(result)}: Failed to dispose object: {str(e)}"
            })
            overall_result = E_FAIL
    
    # Return aggregated results
    return {
        "result": overall_result,
        "message": f"Processed {len(runtime_id_or_ids)} object(s)",
        "details": results
    }


@mcp.tool("ListActiveComObjects")
def list_active_com_objects() -> Dict[str, Any]:
    """
    Lists all currently active COM objects in the server's registry.
    Windows COM API equivalent: Closest equivalent is EnumRunning from the Running Object Table (ROT).
    Output Format: Print the result formatted as a table.
    """
    objects = []
    for obj_id, obj_data in object_registry.items():
        objects.append({
            "runtime_id": obj_id,
            "prog_id": obj_data["prog_id"],  # Use the stored ProgID
            "clsid": obj_data.get("clsid", "Unknown")  # Get CLSID if available
        })

    result = S_OK
    return {
        "result": result,
        "message": f"{hr_to_string(result)}: Found {len(objects)} active COM objects",
        "objects": objects
    }

@mcp.resource("mcp-com://demos/sapi-greeting")
def res_greeting() -> str:
    """This resource runs a tool demo to make the computer say a greeting"""
    return "Use SAPI COM object and say 'Greatings from Johann!'"

@mcp.resource("mcp-com://demos/demo-ie")
def res_greeting() -> str:
    """This resource runs a demo to open IE and navigate to a URL"""
    return "Open IE and navigate to https://embracethered.com"

@mcp.resource("mcp-com://demos/open-excel")
def res_greeting() -> str:
    """This resource runs a demo and opens Excel"""
    return "Open Excel and write 'Hello, world' into the a new worksheet"

if __name__ == "__main__":
    try:
        mcp.run(transport="stdio")
    except Exception as e:
        logger.error(f"mcp-com-server error: {str(e)}")
        sys.exit(1) 