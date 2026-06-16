# core/mapper.py

def apply_mapping(df, mapping):

    rename_dict = {}

    hit_report = {}
    missing_report = {}

    for standard_field, source_fields in mapping.items():

        if isinstance(source_fields, str):
            source_fields = [source_fields]

        found = False

        for source_field in source_fields:

            if source_field in df.columns:

                rename_dict[source_field] = standard_field

                hit_report[standard_field] = source_field

                found = True

                break

        if not found:

            missing_report[standard_field] = source_fields

    df = df.rename(columns=rename_dict)

    return df, hit_report, missing_report