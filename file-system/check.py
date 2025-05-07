import os
import sys

def check_path_security(allowed_base_path: str, target_path: str) -> dict:
    """
    Checks if a target_path is allowed based on the allowed_base_path.

    This involves:
    1. Normalizing and resolving real paths for both inputs.
    2. Checking if the normalized paths exist on the filesystem.
    3. On Windows, checking if paths are on the same drive.
    4. Checking if the target path is a child path (or the same) as the base path
       using a secure relative path comparison.

    Args:
        allowed_base_path: The root directory path that is permitted.
        
        target_path: The path to a file or directory to check.
                     

    Returns:
        A dictionary containing detailed results of the checks:
        - 'original_base': The input allowed_base_path string.
        - 'original_target': The input target_path string.
        - 'normalized_base': The normalized and real path of the base, or None if normalization failed.
        - 'normalized_target': The normalized and real path of the target, or None if normalization failed.
        - 'normalization_successful': bool, True if normalization succeeded.
        - 'base_exists': bool, True if normalized_base_path exists.
        - 'target_exists': bool, True if normalized_target_path exists.
        - 'same_drive_check': bool, True if paths are on the same drive (Windows only, always True otherwise).
        - 'relative_path_from_base': str, The path from normalized_base to normalized_target, or None if not applicable.
        - 'is_within_base': bool, True if the relative path doesn't go outside the base.
        - 'is_allowed': bool, The final decision (True if all checks pass).
        - 'message': str, A human-readable summary of the result.
    """
    results = {
        'original_base': allowed_base_path,
        'original_target': target_path,
        'normalized_base': None,
        'normalized_target': None,
        'normalization_successful': False,
        'base_exists': False,
        'target_exists': False,
        'same_drive_check': True, # Default to True, checked specifically on Windows
        'relative_path_from_base': None,
        'is_within_base': False, # Default to False, set to True if relpath check passes
        'is_allowed': False,     # Final decision, default to False
        'message': "Initial state."
    }

    # --- 1. Path Normalization and Real Path Resolution ---
    try:
        # normpath() cleans up redundant separators, '.', and '..'.
        # realpath() resolves symbolic links, crucial for preventing escapes via symlinks.
        normalized_base_path = os.path.realpath(os.path.normpath(allowed_base_path))
        normalized_target_path = os.path.realpath(os.path.normpath(target_path))

        results['normalized_base'] = normalized_base_path
        results['normalized_target'] = normalized_target_path
        results['normalization_successful'] = True
        results['message'] = "Normalization successful."

    except OSError as e:
         results['message'] = f"Error normalizing paths: {e}"
         return results # Return early on normalization failure
    except Exception as e:
         results['message'] = f"An unexpected error occurred during path normalization: {e}"
         return results # Return early on other normalization errors

    # --- 2. Existence Checks ---
    # Check if the *allowed_base_path* itself exists.
    # It's usually good practice that the root of the allowed area exists.
    if not os.path.exists(normalized_base_path):
        results['base_exists'] = False
        results['message'] = f"Allowed base path does not exist: {allowed_base_path} ({normalized_base_path})"
        return results # Return early if base doesn't exist
    results['base_exists'] = True # Base path exists

    # Check if the *target_path* exists
    if not os.path.exists(normalized_target_path):
        results['target_exists'] = False
        results['message'] = f"Target path does not exist: {target_path} ({normalized_target_path})"
        return results # Return early if target doesn't exist
    results['target_exists'] = True # Target path exists

    # Update message after existence checks pass
    results['message'] = "Paths normalized and exist."

    # --- 3. Confinement Check (Is Target Path Within Base Path?) ---
    try:
        # On Windows, check for different drives explicitly first.
        # os.path.splitdrive() handles case-insensitivity for drives on Windows.
        if sys.platform == 'win32':
            base_drive = os.path.splitdrive(normalized_base_path)[0]
            target_drive = os.path.splitdrive(normalized_target_path)[0]
            if base_drive.lower() != target_drive.lower():
                results['same_drive_check'] = False
                results['message'] = f"Target path is on a different drive ('{target_drive}') than the base path ('{base_drive}')."
                return results # Return early if on different drives

        # Calculate the relative path of the target from the base
        # os.path.relpath(path, start) returns a path from start to path.
        # If target_path is base_path or inside it, relpath will not start with '..'.
        relative_path = os.path.relpath(normalized_target_path, start=normalized_base_path)
        results['relative_path_from_base'] = relative_path

        # Check if the relative path tries to go "up" the directory tree ('..' or starts with '..')
        # os.pardir is the string representation of '..' for the current OS.
        if relative_path == os.pardir or relative_path.startswith(os.pardir + os.sep):
             results['is_within_base'] = False
             results['message'] = f"Target path ('{normalized_target_path}') is outside the allowed base path ('{normalized_base_path}'). Relative path: '{relative_path}'"
             # is_allowed is already False, so no need to set it here again
             return results # Return early if target path goes outside

        # If we reach here, all confinement checks passed.
        results['is_within_base'] = True
        results['is_allowed'] = True
        results['message'] = f"Path is allowed: {target_path} ({normalized_target_path})"
        return results # All checks passed, return success

    except ValueError as e:
        # os.path.relpath can raise ValueError (e.g., paths on different drives before explicit check)
         results['message'] = f"Error calculating relative path: {e}"
         # is_allowed is already False
         return results # Return early on relpath error
    except Exception as e:
         results['message'] = f"An unexpected error occurred during path comparison: {e}"
         # is_allowed is already False
         return results # Return early on other comparison errors



