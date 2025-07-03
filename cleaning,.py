#This code was made on Jupyter notebook, i just pasted here
#Esse código foi feito no Jupyter notebook, apenas copiei para visualização


import pandas as pd

df_orders = pd.read_excel("Nestig _ Assessment [Business Intelligence Analyst I, Database].xlsx", sheet_name=0)
df_order_items = pd.read_excel("Nestig _ Assessment [Business Intelligence Analyst I, Database].xlsx", sheet_name=1)
df_customers = pd.read_excel("Nestig _ Assessment [Business Intelligence Analyst I, Database].xlsx", sheet_name=2)

df_order_items['order_id'].value_counts()

df_order_items['quantity'] = df_order_items['quantity'].fillna(1)
df_order_items['unit_price'] = df_order_items['unit_price'].replace(",", ".", regex=True)
df_order_items['quantity'] = df_order_items['quantity'].replace(",", ".", regex=True)

df_order_items['quantity'] = df_order_items['quantity'].astype(float)
df_order_items['unit_price'] = df_order_items['unit_price'].astype(float)


# Loop para multiplicar linha por linha
resultados = []
for index, row in df_order_items.iterrows():
    resultado = row['quantity'] * row['unit_price']
    resultados.append(resultado)

df_order_items['resultado'] = resultados

df_order_sum = df_order_items.groupby('order_id')['resultado'].sum().reset_index()

df_orders_new = df_orders.merge(df_order_sum, 'right', on='order_id')

df_orders_new = df_orders_new.drop('total_value', axis=1)

df_orders_new['sales_channel'].replace('Instagran','Instagram',inplace=True)
df_orders_new['sales_channel'].replace('Websitea','Website',inplace=True)
df_orders_new['sales_channel'].value_counts()

df_orders_new['sales_channel'].isna().value_counts()

df_order_items['category'] = df_order_items['category'].replace("Wallpapers","Wallpaper")
df_order_items['category'] = df_order_items['category'].replace("Cribs","Crib")
df_order_items['category'] = df_order_items['category'].str.strip()

df_order_items.isnull().value_counts()

df_orders_new = df_orders_new.rename(columns={'resultado': 'total_value'})

order_counts = df_orders_new.groupby('customer_id')['order_id'].count().reset_index()
order_counts.rename(columns={'order_id': 'order_count'}, inplace=True)

df_customers_new = pd.merge(df_customers, order_counts, on='customer_id', how='left')

# Criar a coluna 'customer_type' (Novo ou Repetido)
df_customers_new ['customer_type'] = df_customers_new['order_count'].apply(lambda x: 'Novo' if x == 1 else ('Repetido' if x > 1 else None))

with pd.ExcelWriter('nestig_database.xlsx') as writer:
    # Escreve cada DataFrame em uma aba diferente
    df_orders_new.to_excel(writer, sheet_name='Orders', index=False)
    df_order_items.to_excel(writer, sheet_name='Order_item', index=False)
    df_customers_new.to_excel(writer, sheet_name='Customer', index=False)