let
    Source =
        Csv.Document(
            File.Contents("C:\Users\win10\OneDrive\Desktop\Sales Analysis\outputs\cleaned_sales.csv"),
            [
                Delimiter = ",",
                Encoding = 65001,
                QuoteStyle = QuoteStyle.Csv
            ]
        ),
    PromotedHeaders =
        Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    TypedColumns =
        Table.TransformColumnTypes(
            PromotedHeaders,
            {
                {"Order ID", type text},
                {"Order Date", type date},
                {"Customer", type text},
                {"Product", type text},
                {"Category", type text},
                {"Region", type text},
                {"Quantity", type number},
                {"Sales", type number},
                {"Profit", type number}
            },
            "en-US"
        )
in
    TypedColumns
