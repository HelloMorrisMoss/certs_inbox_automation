"""These customers will have many more certs generated than they require.

Certs created with the same values in the columns below will have additional beyond the first moved from the mailbox."""

dedupe_cnums: tuple = ('1234', '4321')
dedupe_columns: list = ['product', 'so_number', 'lot']
