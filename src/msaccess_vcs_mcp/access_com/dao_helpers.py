"""Helper utilities for working with Access DAO objects."""

from typing import Any


def get_query_type_name(query_type: int) -> str:
    """
    Get query type name from QueryDef.Type value.
    
    Args:
        query_type: QueryDef.Type integer value
        
    Returns:
        Human-readable query type name
    """
    type_map = {
        0: "Select",
        1: "Union",
        2: "PassThrough",
        3: "DataDefinition",
        4: "Append",
        5: "Delete",
        6: "Update",
        7: "MakeTable",
    }
    return type_map.get(query_type, "Unknown")


def list_table_defs(db) -> list[dict[str, Any]]:
    """
    List all user tables from database.
    
    Args:
        db: DAO Database object
        
    Returns:
        List of table metadata dictionaries
    """
    tables = []
    for table_def in db.TableDefs:
        # Skip system tables
        if not table_def.Name.startswith("MSys"):
            tables.append({
                "name": table_def.Name,
                "record_count": None,  # Would require expensive RecordCount operation
            })
    return tables


def list_query_defs(db) -> list[dict[str, Any]]:
    """
    List all queries from database.
    
    Args:
        db: DAO Database object
        
    Returns:
        List of query metadata dictionaries
    """
    queries = []
    for query_def in db.QueryDefs:
        queries.append({
            "name": query_def.Name,
            "type": get_query_type_name(query_def.Type),
        })
    return queries


def get_query_sql(db, query_name: str) -> str:
    """
    Get SQL from a saved query.
    
    Args:
        db: DAO Database object
        query_name: Name of the query
        
    Returns:
        SQL string
        
    Raises:
        KeyError: If query not found
    """
    try:
        query_def = db.QueryDefs(query_name)
        return query_def.SQL
    except Exception as e:
        raise KeyError(f"Query '{query_name}' not found: {e}")


def list_modules(db) -> list[str]:
    """
    List all VBA modules from database.
    
    Note: This returns standard modules only.
    Form/Report modules need to be accessed via the Access Application object.
    
    Args:
        db: DAO Database object
        
    Returns:
        List of module names
    """
    modules = []
    try:
        # Access modules through the Application's VBE
        # This is typically done through the Access Application object
        # For now, return empty list - will be implemented in exporter
        pass
    except Exception:
        pass
    return modules


def get_table_schema(db, table_name: str) -> dict[str, Any]:
    """
    Get schema information for a table.
    
    Args:
        db: DAO Database object
        table_name: Name of the table
        
    Returns:
        Dictionary with table schema information
        
    Raises:
        KeyError: If table not found
    """
    try:
        table_def = db.TableDefs(table_name)
        fields = []
        for field in table_def.Fields:
            field_info = {
                "name": field.Name,
                "type": field.Type,
                "size": field.Size if hasattr(field, 'Size') else None,
                "required": field.Required if hasattr(field, 'Required') else False,
                "allow_zero_length": field.AllowZeroLength if hasattr(field, 'AllowZeroLength') else False,
            }
            fields.append(field_info)
        
        indexes = []
        for index in table_def.Indexes:
            index_fields = [f.Name for f in index.Fields]
            index_info = {
                "name": index.Name,
                "fields": index_fields,
                "primary": index.Primary if hasattr(index, 'Primary') else False,
                "unique": index.Unique if hasattr(index, 'Unique') else False,
            }
            indexes.append(index_info)
        
        return {
            "name": table_name,
            "fields": fields,
            "indexes": indexes,
        }
    except Exception as e:
        raise KeyError(f"Table '{table_name}' not found: {e}")
