import pandas as pd

def join_dataframes_on_indexes(df1, df2, matches):
    df_joined = pd.DataFrame()

    for match in matches:
        temp_df1 = df1.iloc[[match[0]]]
        temp_df2 = df2.iloc[[match[1]]]

        temp_df1 = temp_df1.reset_index()
        temp_df2 = temp_df2.reset_index()

        temp_joined = pd.concat([temp_df1, temp_df2], axis=1)
        df_joined = pd.concat([df_joined, temp_joined])

    return df_joined.reset_index(drop=True)

df1 = pd.DataFrame({
    'columnname': ['asdf', 'asdfasdf', 'x', 'y', 'z']
})

df2 = pd.DataFrame({
    'cola': ['a', 'e', 'b'],
    'colb': ['a', 'g', 'b'],
    'colc': ['x', 'b', 'b'],
    'cold': ['b', 'n', 'b']
})

matches = [[0, 2], [0, 0], [3, 0], [4,2]]

print(df1)
print(df2)

df_joined = join_dataframes_on_indexes(df1, df2, matches)
print(df_joined)
