# SOP
c.execute("""
CREATE TABLE IF NOT EXISTS sop (
    user_id INTEGER,
    purchasing TEXT,
    receiving TEXT,
    storage TEXT
)
""")

# TRACEABILITY
c.execute("""
CREATE TABLE IF NOT EXISTS traceability (
    user_id INTEGER,
    system_desc TEXT
)
""")

# RECALL
c.execute("""
CREATE TABLE IF NOT EXISTS recall (
    user_id INTEGER,
    procedure TEXT
)
""")

# EVALUATION
c.execute("""
CREATE TABLE IF NOT EXISTS evaluation (
    user_id INTEGER,
    method TEXT
)
""")
