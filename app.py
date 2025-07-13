import pandas as pd
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
import numpy as np
from reportlab.lib.units import cm 
import matplotlib.pyplot as plt
import streamlit as st
import io

# Función para dibujar el fondo
def draw_cover(canvas, doc):
    width, height = doc.pagesize
    canvas.setFillColor(rl_colors.lightgrey)
    canvas.rect(0, 0, width, height, fill=1, stroke=0)

    # Título centrado
    canvas.setFont("Helvetica", 50)
    canvas.setFillColor(rl_colors.black)
    canvas.drawCentredString(width / 2, height / 2 + 20, "Defects & Warranty")
    canvas.drawCentredString(width / 2, height / 2 - 35, "Report")

def procesar_archivos(defectFile, productionFile):
    df= pd.read_excel(defectFile, header=1)
    fechainicio = pd.to_datetime("2025-06-30")
    fechainicio2 = pd.to_datetime("2024-12-30")
    df["semana_relativa"] = (((df['Date:'] - fechainicio).dt.days // 7) + 1).astype("Int64")
    df["Historical Week"] = 'Week ' + df["semana_relativa"].astype(str) 
    df["semana_natural"] = (((df["Date:"] - fechainicio2).dt.days // 7) + 1).astype("Int64")
    df["Year Week"] = 'Week ' + df["semana_natural"].astype(str)
    first_nan_index = df[df[["Date:"]].isnull().any(axis=1)].index.min()
    df = df.iloc[:first_nan_index, :]

    #Catalogo de defectos
    defect_type = {
        "New 60* Day Reshaft Policy - Customer Service":"FR30DAYRESHAFT",
        "Customer Service" : "FRCUST",
        "Damaged in Shipping":"FRDAMAGE",
        "Apparel - Broken Component" : "FRDEFECT",
        "Broken Hosel" : "FRDEFECT",
        "Broken Shaft - grip" : "FRDEFECT",
        "Broken Shaft - hosel" : "FRDEFECT",
        "Club Head Rattle" : "FRDEFECT",
        "Components Stripped (Adapter / Screw)" : "FRDEFECT",
        "Face Caving In" : "FRDEFECT",
        "Face Crack" : "FRDEFECT",
        "Ferrule Repair" : "FRDEFECT",
        "Popped Crown" : "FRDEFECT",
        "Putter Weight Detached" : "FRDEFECT",
        "Shaft Paint Chipping" : "FRDEFECT",
        "Weight Port Failure" : "FRDEFECT",
        "Crown Dent" : "FRDEFECT",
        "Head paint / finish chipping" : "FRDEFECT",
        "Poor Etching" : "FRDEFECT",
        "Bag - Broken Component" : "FRDEFECT",
        "Missing Paint Fill" : "FRDEFECT",
        "Order Entry - Wrong Component" : "FRERROR",
        "Order Entry - Wrong Dexterity" : "FRERROR",
        "Order Entry - Length, Loft, Lie, etc." : "FRERROR",
        "Misfit" : "FRERROR",
        "Order Entry - Missing Item" : "FRERROR",
        "Lost in shipping/Delivered to wrong address" : "FRLOSTSHIP",
        "Blemishes / Cosmetic Damage" : "FRMISBUILD",
        "Broken Component - Assembly" : "FRMISBUILD",
        "Club Length - too long" : "FRMISBUILD",
        "Club Length - too short" : "FRMISBUILD",
        "Complete Head Shaft Separation" : "FRMISBUILD",
        "Goop in Hosel" : "FRMISBUILD",
        "Grip Alignment" : "FRMISBUILD",
        "Loose Putter Hosel" : "FRMISBUILD",
        "Loose Shaft / Rattle" : "FRMISBUILD",
        "Assembly - Wrong Component" : "FRMISBUILD",
        "Misbuild - Loft, Lie, Swingweight" : "FRMISBUILD",
        "Loose Weight" : "FRMISBUILD",
        "Assembly - Wrong Order" : "FRWRONG",
        "Apparel - Wrong Component" : "FRWRONG",
        "Apparel - Wrong Order" : "FRWRONG",
        "Swing Weights" : "REPAIR",
        "-" : "RETURN",
        "Paid reshaft": "SHAFTREPAIR",
        "No defect found": "SHIPPING ONLY",
        "M16 Shaft Connection" : "FRDEFECT",
        "Toe Dent" : "FRDEFECT"

    }
    semana_actual = "Week 1"
    df["Type"] = df["Claim Type (Description)"].map(defect_type)

    df = df[["Date:", "Historical Week", "Year Week", "Shipper:", "Original Order or Serial #", "RMA", "RC", "Status? (0,1,2)","Shipping Carrier","Tracking Number",
            "Staged", "Make / Model", "Claim Type (Description)", "Type", "Pod Number", "Original Sales Order Date" ]]

    #Crear PDF
    doc = SimpleDocTemplate("Defects_Warranty.pdf", pagesize = landscape(letter))
    styles = getSampleStyleSheet()
    story = []
    story.append(PageBreak())
    table_style = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('SPAN', (0, 0), (-1, 0)), #Add title
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(1,1),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
        ('BACKGROUND', (0,-1), (-1,-1), rl_colors.lightgrey),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
    ]
    table_style_semana_actual = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(1,1),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
    ]
    table_style_graphic = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(1,1),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
        ('BACKGROUND', (0,-1), (-1,-1), rl_colors.lightgrey),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
    ]
    table_style_graphic2 = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
        ('BACKGROUND', (0,-1), (-1,-1), rl_colors.lightgrey),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
    ]

    table_style_weeks = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('SPAN', (0, 0), (-1, 0)), #Add title
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(0,1),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
        ('BACKGROUND', (0,-1), (-1,-1), rl_colors.lightgrey),
        ('FONTNAME',(0,-2),(-1,-1),'Helvetica-Bold'),
    ]
    prod_style = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('SPAN', (0, 0), (-1, 0)), #Add title
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(1,1),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
        ('BACKGROUND', (0,-2), (-1,-1), rl_colors.lightgrey),
        ('FONTNAME',(0,-2),(-1,-1),'Helvetica-Bold'),
    ]
    prod_style_weeks = [
        ('FONTSIZE', (0,0), (-1,-1), 9), 
        ('SPAN', (0, 0), (-1, 0)), #Add title
        ('SPAN', (0,2), (-1,2)),
        ('SPAN', (0,3), (-1,3)),
        ('BACKGROUND', (0,0), (-1,0), rl_colors.darkgrey),
        ('ALIGN',(0,0),(-1,0),'CENTER'),
        ('FONTNAME',(0,0),(-1,1),'Helvetica-Bold'),
        ('ALIGN',(0,1),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,rl_colors.black),
        ('BACKGROUND', (0,-2), (-1,-1), rl_colors.lightgrey),
        ('FONTNAME',(0,-2),(-1,-1),'Helvetica-Bold'),
    ]
    fourweeks = sorted(df['Historical Week'].unique())[-4:]
    eightweeks = sorted(df['Historical Week'].unique())[-8:]
    df8 = df[df['Historical Week'].isin(eightweeks)]
    df4 = df[df['Historical Week'].isin(fourweeks)]

    #Staged
    staged = pd.crosstab(df4['Staged'], df4['Historical Week'])
    staged.loc['Total'] = staged.sum(numeric_only=True)
    #Data
    staged_data = [['Count of Staged by Week']]
    staged_data += [['Staged'] + staged.columns.tolist()] 
    for idx, row in staged.iterrows():
        staged_data.append([idx] + row.astype(int).astype(str).tolist())
    #Tabla
    num_filas_staged = len(staged_data)
    row_heights_staged = [13] * num_filas_staged
    staged_table = Table(staged_data, colWidths=[100, 60, 60, 60, 60], repeatRows=1, rowHeights=row_heights_staged)
    staged_table.setStyle(TableStyle(table_style))

    #Avg Staged
    week_cols = [col for col in staged.columns if col.startswith('Week')]
    # Calcular TOTAL
    staged['TOTAL'] = staged[week_cols].sum(axis=1)
    # Contar cuántos valores NO NULOS hay por fila en esas columnas
    non_null_weeks = staged[week_cols].notnull().sum(axis=1)
    # Calcular AVG dinámico
    staged['AVG'] = (staged['TOTAL'] / non_null_weeks).round(0).astype(int)
    avg = staged[['AVG','TOTAL']].copy()
    #Tabla
    avg_data = [['Last 4 Weeks']]
    avg_data += [list(avg.columns)]
    avg_data += avg.values.tolist()
    num_filas_staged_avg = len(avg_data)
    row_heights_staged_avg = [13] * num_filas_staged_avg 
    avg_table = Table(avg_data, colWidths=[60,60], rowHeights=row_heights_staged_avg)
    avg_table.setStyle(TableStyle(table_style_weeks))
    #First Two Tables
    joined_staged = Table([[staged_table, avg_table, '']])
    story.append(joined_staged)

    #Warranty Details
    warranty = pd.crosstab(df4['Type'], df4['Historical Week'])
    warranty.loc['Total'] = warranty.sum(numeric_only=True)
    warranty_data = [['Warranty Details']]
    warranty_data += [['Type'] + warranty.columns.tolist()]
    for idx, row in warranty.iterrows():
        warranty_data.append([idx] + row.astype(int).astype(str).tolist())
    #Tabla
    num_filas_w = len(warranty_data)
    row_heights_w = [13] * num_filas_w
    warranty_table = Table(warranty_data, colWidths=[100, 60, 60, 60, 60], repeatRows=1, rowHeights=row_heights_w)
    warranty_table.setStyle(TableStyle(table_style))

    #Avg Details
    warranty['TOTAL'] = warranty[week_cols].sum(axis=1)
    non_null_weeks_warranty = warranty[week_cols].notnull().sum(axis=1)
    warranty['AVG'] = (warranty['TOTAL'] / non_null_weeks_warranty).round(0).astype(int)
    avg_warranty = warranty[['AVG','TOTAL']].copy()
    #Tabla 4 weeks
    avg_warranty_data = [['Last 4 Weeks']]
    avg_warranty_data += [list(avg_warranty.columns)]
    avg_warranty_data += avg_warranty.values.tolist()
    num_filas_avg_w = len(avg_warranty_data)
    row_heights_avg_W = [13] * num_filas_avg_w
    avg_warranty_table = Table(avg_warranty_data, colWidths=[60,60], rowHeights=row_heights_avg_W)
    avg_warranty_table.setStyle(TableStyle(table_style_weeks))
    #Tabla 8 weeks
    warranty8 = pd.crosstab(df4['Type'], df8['Historical Week'])
    warranty8.loc['Total'] = warranty8.sum(numeric_only=True)
    warranty8['TOTAL'] = warranty8[week_cols].sum(axis=1)
    non_null_weeks_warranty8 = warranty8[week_cols].notnull().sum(axis=1)
    warranty8['AVG'] = (warranty8['TOTAL'] / non_null_weeks_warranty8).round(0).astype(int)
    avg_warranty8 = warranty8[['AVG','TOTAL']].copy()
    avg_warranty_data8 = [['Last 8 Weeks']]
    avg_warranty_data8 += [list(avg_warranty8.columns)]
    avg_warranty_data8 += avg_warranty8.values.tolist()
    num_filas_avg_w8 = len(avg_warranty_data8)
    row_heights_w8 = [13] * num_filas_avg_w8
    avg_warranty_table8 = Table(avg_warranty_data8, colWidths=[60,60],rowHeights=row_heights_w8)
    avg_warranty_table8.setStyle(TableStyle(table_style_weeks))
    #Second Tables
    joined_warranty = Table([[warranty_table,  avg_warranty_table,avg_warranty_table8]])
    story.append(joined_warranty)

    #ORDENES Y PRODUCCION

    dfprod = pd.read_csv(productionFile, encoding='utf-16',sep='\t', header=1)
    dfprod = dfprod.drop(columns=[col for col in dfprod.columns 
                    if 'Unnamed' in str(col) and col != 'Unnamed: 2'])
    dfprodfilter = dfprod[dfprod['SiteName'] == 'Grand Total']
    dfprodfilter = dfprodfilter[dfprodfilter['Unnamed: 2'] != 'Shipments']
    dfprodfilter = dfprodfilter.drop(['Local Operations Shift','Grand Total'], axis=1)
    dfprodfilter = dfprodfilter.drop(['SiteName'], axis=1)
    dfprodfilter = dfprodfilter.reset_index(drop=True)
    #Transponer el DataFrame para que las métricas sean columnas
    df_transposed = dfprodfilter.set_index('Unnamed: 2').T.reset_index()
    df_transposed.columns = ['Fecha', 'Orders', 'ShippedQty']  # Renombrar columnas
    df_transposed['Orders'] = df_transposed['Orders'].astype(int)
    df_transposed['ShippedQty'] = (
        df_transposed['ShippedQty']
        .astype(str)  # Convertir todo a string primero
        .str.replace('[.,]', '', regex=True)  # Eliminar puntos y comas
        .replace('nan', np.nan)  # Mantener NaN como valores nulos
        .astype('Int64')  # Tipo nullable integer de pandas
    ) # Eliminar comas
    #Agreagar semanas
    def convert_date(date_str):
        # Intenta con el formato preferido primero
        try:
            return pd.to_datetime(date_str + '-2025', format='%b %d-%Y')
        except ValueError:
            # Fallback 1: Formato sin espacio (ej: "jul-01-2025")
            try:
                return pd.to_datetime(date_str + '-2025', format='%b-%d-%Y')
            except:
                # Fallback 2: Formato ISO (ej: "2025-07-01")
                try:
                    return pd.to_datetime(date_str, format='ISO8601')
                except:
                    # Fallback 3: Intento automático
                    return pd.to_datetime(date_str, format='mixed', dayfirst=False)

    # Aplicar la conversión robusta
    df_transposed['Fecha'] = df_transposed['Fecha'].apply(convert_date)
    #df_transposed['Fecha'] = pd.to_datetime(df_transposed['Fecha'] + '-2025', format='%b %d-%Y')
    df_transposed["semana_relativa"] = (((df_transposed['Fecha'] - fechainicio).dt.days // 7) + 1).astype("Int64")
    df_transposed["Historical Week"] = 'Week ' + df_transposed["semana_relativa"].astype(str) 
    df_transposed["semana_natural"] = (((df_transposed["Fecha"] - fechainicio2).dt.days // 7) + 1).astype("Int64")
    df_transposed["Year Week"] = 'Week ' + df_transposed["semana_natural"].astype(str)
    df_transposed = df_transposed[['Fecha', 'Historical Week','Year Week','Orders', 'ShippedQty',]]

    prod4 = df_transposed[df_transposed['Historical Week'].isin(fourweeks)]
    prod8 = df_transposed[df_transposed['Historical Week'].isin(eightweeks)]


    # Agrupar por 'Historical Week' y sumar Orders/ShippedQty
    df_weekly = prod4.groupby('Historical Week', as_index=False).agg({
        'Orders': 'sum',
        'ShippedQty': 'sum',
        'Fecha': ['min', 'max']  # Primera y última fecha de la semana
    })
    # Renombrar columnas para claridad
    df_weekly.columns = ['Week', 'Total Orders', 'Total ShippedQty', 'Start Date', 'End Date']
    # Formatear fechas como "DD-MMM" (ej: "22-Apr")
    df_weekly['Start Date'] = df_weekly['Start Date'].dt.strftime('%d-%b')
    df_weekly['End Date'] = df_weekly['End Date'].dt.strftime('%d-%b')
    # Ordenar por semana (opcional)
    df_weekly = df_weekly.sort_values('Week', ascending=True)
    #Production Data
    prod_data = [
        ["Production data"],
        ["", *df_weekly['Week'].tolist()], 
        ["Start Date"] + df_weekly['Start Date'].tolist(),
        ["End Date"] + df_weekly['End Date'].tolist(),
        ["ASM Clubs"] + [f"{x:,.0f}" for x in df_weekly['Total ShippedQty']],
        ["Orders"] + [f"{x:,.0f}" for x in df_weekly['Total Orders']]
    ]
    num_filas_prod = len(prod_data)
    row_heights_prod = [13] * num_filas_prod
    prod_tabla = Table(prod_data, colWidths=[100, 60, 60, 60, 60], repeatRows=1, rowHeights=row_heights_prod)
    prod_tabla.setStyle(TableStyle(prod_style))

    #Avg Details
    # 1. Calcular métricas para ASM Clubs y Orders
    start_date = df_weekly['Start Date'].iloc[0]
    end_date = df_weekly['End Date'].iloc[-1]  
    num_semanas = len(df_weekly)

    # Para ASM Clubs (Total ShippedQty)
    total_asm = df_weekly['Total ShippedQty'].sum()
    avg_asm = (total_asm / num_semanas).round(0).astype(int)

    # Para Orders
    total_orders = df_weekly['Total Orders'].sum()
    avg_orders = (total_orders / num_semanas).round(0).astype(int)

    # 2. Preparar datos para la tabla
    avg_prod_data = [
        ["Last 4 Weeks", ""],  # Título combinado
        ["AVG", "TOTAL"],
        [start_date],  
        [end_date],  
        [f"{avg_asm:,}", f"{total_asm:,}"],  # Fila ASM Clubs
        [f"{avg_orders:,}", f"{total_orders:,}"]   # Fila Orders
    ]
    num_filas_prod_avg = len(avg_prod_data)
    row_heights_prod_avg = [13] * num_filas_prod_avg 
    avg_prod_table = Table(avg_prod_data, colWidths=[60,60], rowHeights=row_heights_prod_avg)
    avg_prod_table.setStyle(TableStyle(prod_style_weeks))

    #8 WEEKS PRODUCTION
    # Agrupar por 'Historical Week' y sumar Orders/ShippedQty
    df_weekly8 = df_transposed.groupby('Historical Week', as_index=False).agg({
        'Orders': 'sum',
        'ShippedQty': 'sum',
        'Fecha': ['min', 'max']  # Primera y última fecha de la semana
    })

    # Renombrar columnas para claridad
    df_weekly8.columns = ['Week', 'Total Orders', 'Total ShippedQty', 'Start Date', 'End Date']
    # Formatear fechas como "DD-MMM" (ej: "22-Apr")
    df_weekly8['Start Date'] = df_weekly8['Start Date'].dt.strftime('%d-%b')
    df_weekly8['End Date'] = df_weekly8['End Date'].dt.strftime('%d-%b')
    # Ordenar por semana (opcional)
    df_weekly8 = df_weekly8.sort_values('Week', ascending=True)
    #Avg Details
    # 1. Calcular métricas para ASM Clubs y Orders
    start_date8 = df_weekly8['Start Date'].iloc[0]
    end_date8 = df_weekly8['End Date'].iloc[-1]  
    num_semanas8 = len(df_weekly8)

    # Para ASM Clubs (Total ShippedQty)
    total_asm8 = df_weekly8['Total ShippedQty'].sum()
    avg_asm8 = (total_asm8 / num_semanas8).round(0).astype(int)

    # Para Orders
    total_orders8 = df_weekly8['Total Orders'].sum()
    avg_orders8 = (total_orders8 / num_semanas8).round(0).astype(int)

    # 2. Preparar datos para la tabla
    avg_prod_data8 = [
        ["Historical", ""],  # Título combinado
        ["AVG", "TOTAL"],
        [start_date8],  
        [end_date8],  
        [f"{avg_asm8:,}", f"{total_asm8:,}"],  # Fila ASM Clubs
        [f"{avg_orders8:,}", f"{total_orders8:,}"]   # Fila Orders
    ]
    num_filas_avg_prod8 = len(avg_prod_data8)
    row_heights_avg_prod8 = [13] * num_filas_avg_prod8 
    avg_prod_table8 = Table(avg_prod_data8, colWidths=[60,60], rowHeights=row_heights_avg_prod8)
    avg_prod_table8.setStyle(TableStyle(prod_style_weeks))
    joined_prod = Table([[prod_tabla, avg_prod_table,avg_prod_table8]])
    story.append(joined_prod)

    #WEEKLY ORDERS
    # Suma de ordenes por semana
    weekly_orders_totals = df_weekly.set_index('Week')['Total Orders']

    # Errores de Warranty
    orders = pd.crosstab(df4['Type'], df4['Historical Week'])
    # Division porcentual
    orders_pct = (orders.div(weekly_orders_totals) * 100).round(1)
    sum_pct = orders_pct.sum().round(1)

    orders_data = [['Weekly Errors']]  # Título modificado
    orders_data += [['Type'] + orders_pct.columns.tolist()]
    for idx, row in orders_pct.iterrows():
        formatted_values = [f"{val}%" if not pd.isna(val) else "0%" for val in row]
        orders_data.append([idx] + formatted_values)

    # 5. Añadir fila de totales (opcional)
    orders_data.append(['Order Quality'] + [f"{x}%" for x in sum_pct])
    num_filas = len(orders_data)
    row_heights = [13] * num_filas 
    #Tabla
    orders_table = Table(orders_data, colWidths=[100, 60, 60, 60, 60], repeatRows=1, rowHeights=row_heights)
    orders_table.setStyle(TableStyle(table_style))

    #Avg Orders %
    orders_pct['TOTAL'] = orders_pct[week_cols].sum(axis=1)
    non_null_weeks_orders = orders_pct[week_cols].notnull().sum(axis=1)
    orders_pct['AVG'] = (orders_pct['TOTAL'] / non_null_weeks_orders).round(1).astype(float)
    avg_orders_pct= orders_pct[['AVG','TOTAL']].copy()

    #Formateo
    avg_orders_pct['AVG'] = (avg_orders_pct['AVG']).round(1).astype(str) + '%'
    avg_orders_pct['TOTAL'] = (avg_orders_pct['TOTAL']).round(1).astype(str) + '%'

    #Tabla 4 weeks
    avg_orders_data = [['Last 4 Weeks']]
    avg_orders_data += [list(avg_orders_pct.columns)]
    avg_orders_data += avg_orders_pct.values.tolist()
    avg_weekly_pct = (sum_pct / len(week_cols)).round(1) 
    total_errors = sum_pct[week_cols].sum()
    avg_orders_data.append([
        f"{avg_weekly_pct.mean()}%",  
        f"{total_errors.mean()}%"           
    ])
    num_filas_avg_opct = len(avg_orders_data)
    row_heights_avg_opct = [13] * num_filas_avg_opct
    avg_orders_table = Table(avg_orders_data, colWidths=[60,60], rowHeights=row_heights_avg_opct)
    avg_orders_table.setStyle(TableStyle(table_style_weeks))

    #Historical
    #Avg Orders %
    # Suma de ordenes por semana historico
    weekly_orders_totals_hist = df_weekly8.set_index('Week')['Total Orders']
    # Errores de Warranty
    orders_hist = pd.crosstab(df4['Type'], df['Historical Week'])
    # Division porcentual
    orders_pct_hist = (orders_hist.div(weekly_orders_totals_hist) * 100).round(1)
    sum_pct_hist = orders_pct_hist.sum().round(1)

    orders_pct_hist['TOTAL'] = orders_pct_hist.sum(axis=1)
    orders_pct_hist['AVG'] = (orders_pct_hist['TOTAL'] / num_semanas8).round(1).astype(float)
    avg_orders_pct_hist= orders_pct_hist[['AVG','TOTAL']].copy()

    #Formateo
    avg_orders_pct_hist['AVG'] = (avg_orders_pct_hist['AVG']).round(1).astype(str) + '%'
    avg_orders_pct_hist['TOTAL'] = (avg_orders_pct_hist['TOTAL']).round(1).astype(str) + '%'

    #Tabla 4 weeks
    avg_orders_data_hist = [['Runnig Total']]
    avg_orders_data_hist += [list(avg_orders_pct_hist.columns)]
    avg_orders_data_hist += avg_orders_pct_hist.values.tolist()
    avg_weekly_pct_hist = (sum_pct_hist / num_semanas8).round(1) 
    total_errors_hist = sum_pct_hist.sum()
    avg_orders_data_hist.append([
        f"{avg_weekly_pct_hist.mean()}%",  
        f"{total_errors_hist.mean()}%"           
    ])
    num_filas_avg_opct_hist = len(avg_orders_data_hist)
    row_heights_avg_opct_hist = [13] * num_filas_avg_opct_hist
    avg_orders_table_hist = Table(avg_orders_data_hist, colWidths=[60,60], rowHeights=row_heights_avg_opct_hist)
    avg_orders_table_hist.setStyle(TableStyle(table_style_weeks))

    orders_joined = Table([[orders_table, avg_orders_table, avg_orders_table_hist]])
    story.append(orders_joined)

    story.append(PageBreak())

    #Grafica
    #DATA
    #Warranty Details
    warranty_hist1 = pd.crosstab(df['Type'], df['Historical Week'])

    # CORRECCIÓN: Usar warranty_hist en lugar de warranty
    warranty_hist1.loc['Total'] = warranty_hist1.sum(numeric_only=True)

    warranty_hist8 = pd.crosstab(df8['Type'], df8['Historical Week'])
    warranty_hist8.loc['Total'] = warranty_hist8.sum(numeric_only=True)
    # 2. Preparar datos para el DataFrame
    #warranty_hist = [['Warranty Details']]
    warranty_hist = [['Type'] + warranty_hist8.columns.tolist()]
    for idx, row in warranty_hist8.iterrows():
        warranty_hist.append([idx] + row.astype(int).astype(str).tolist())

    warranty_data_hist_FG = [['Type'] + warranty_hist1.columns.tolist()]

    for idx, row in warranty_hist1.iterrows():
        row_values = []
        for val in row:
            if pd.isna(val):
                row_values.append('0')
            else:
                row_values.append(str(int(val)))
        warranty_data_hist_FG.append([idx] + row_values)

    # 3. Crear DataFrame
    df_warranty_hist = pd.DataFrame(warranty_data_hist_FG[1:], columns=warranty_data_hist_FG[0])
    df_warranty_hist.set_index('Type', inplace=True)
    df_warranty_hist = df_warranty_hist.apply(pd.to_numeric)

    # 4. Preparar datos para gráfico (excluyendo 'Total')
    plot_data = df_warranty_hist.drop('Total', errors='ignore')

    # 5. Configurar gráfico
    plt.figure(figsize=(12, 8))
    markers = ['o', 's', '^', 'D', 'v', 'p', '*']
    colors = plt.cm.tab20(np.linspace(0, 1, len(plot_data)))

    # 6. Graficar cada línea
    for i, (idx, row) in enumerate(plot_data.iterrows()):
        plt.plot(row.index, row.values, 
                marker=markers[i%len(markers)],
                linestyle='--',
                linewidth=2,
                markersize=8,
                color=colors[i],
                label=idx)

    # 7. Personalización del gráfico
    plt.title('Returns per Week by Reason Code', fontsize=16, pad=20)
    plt.xlabel('Week', fontsize=12)
    plt.ylabel('Number of Returns', fontsize=12)
    plt.xticks(rotation=45)
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.legend(title='Outcome', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()

    # 8. Añadir etiquetas de valores
    for idx, row in plot_data.iterrows():
        for week, val in row.items():
            if val > 0:
                plt.text(week, val + 0.5, str(val), 
                        ha='center', va='bottom', 
                        fontsize=9, color=colors[plot_data.index.get_loc(idx)])

    # Guardar imagen temporal
    plt.savefig("graphic.png", dpi=300)
    plt.close()

    # Insertar la imagen (gráfica)
    story.append(Image("graphic.png", width=750, height=300))  


    # 2. Preparar datos para la tabla de historial
    hist_data = [['Type'] + warranty_hist8.columns.tolist()]
    for idx, row in warranty_hist8.iterrows():
        hist_data.append([idx] + row.astype(int).astype(str).tolist())

    # 3. Crear tabla de historial
    num_filas_warranty_hist = len(hist_data)
    row_heights_w_hist = [13] * num_filas_warranty_hist
    warranty_table_hist = Table(hist_data, colWidths=[100, 58, 58, 58, 58, 58], repeatRows=1, rowHeights=row_heights_w_hist)
    warranty_table_hist.setStyle(TableStyle(table_style_graphic))
    # 4. Calcular los datos de resumen
    last_8_weeks = warranty_hist8.iloc[:, -8:].sum(axis=1) #if len(warranty_hist8.columns) <= 8 else pd.Series(0, index=warranty_hist8.index)
    weeks_5_to_8 = warranty_hist8.iloc[:, -8:-4].sum(axis=1) #if len(warranty_hist8.columns) <= 8 else pd.Series(0, index=warranty_hist8.index)
    last_4_weeks = warranty_hist8.iloc[:, -4:].sum(axis=1) #if len(warranty_hist8.columns) >= 4 else pd.Series(0, index=warranty_hist8.index)

    # 5. Preparar datos para la tabla de resumen
    summary_data = [['Last 4 Weeks', 'Weeks 5-8', 'Dif','Last 8 Weeks']]
    for idx in warranty_hist8.index:
        summary_data.append([
            str(int(last_4_weeks[idx])),
            str(int(weeks_5_to_8[idx])),
            str(int(last_4_weeks[idx])-int(weeks_5_to_8[idx])),
            str(int(last_8_weeks[idx]))        
        ])

    # 6. Crear tabla de resumen
    summary_table = Table(summary_data, colWidths=[58, 58, 58, 58], repeatRows=1, rowHeights=row_heights_w_hist)
    summary_table.setStyle(TableStyle(table_style_graphic2))
    graphic_joined = Table([[warranty_table_hist, summary_table]])
    story.append(graphic_joined)

    story.append(PageBreak())

    #TABLA SEMANA ACTUAL
    df_semana_actual = df[df['Historical Week'] == semana_actual] 
    df_semana_actual = df_semana_actual[["Date:", "Historical Week",  "Shipper:", "Original Order or Serial #","RMA", "RC","Shipping Carrier","Staged", "Claim Type (Description)", "Type"]]
    df_semana_actual['Date:'] = pd.to_datetime(df_semana_actual['Date:']).dt.strftime('%m/%d/%Y')
    rename_columns = {
        'Date:': 'Date',
        'Historical Week': 'Week',
        'Original Order or Serial #': 'Original Order',
        'Claim Type (Description)' : 'Description'
    }
    df_semana_actual = df_semana_actual.rename(columns=rename_columns)
    semana_actual_data = [df_semana_actual.columns.tolist()]  # Encabezados
    semana_actual_data += df_semana_actual.values.tolist()    # Datos
    semana_actual_table = Table(semana_actual_data,repeatRows=2)
    semana_actual_table.setStyle(TableStyle(table_style_semana_actual))
    story.append(semana_actual_table)

    story.append(PageBreak())

    #Resumen de ordenes
    df_weekly8 = prod8.groupby('Historical Week', as_index=False).agg({
        'Orders': 'sum',
        'ShippedQty': 'sum',
        'Fecha': ['min', 'max']  # Primera y última fecha de la semana
    })
    # Renombrar columnas para claridad
    df_weekly8.columns = ['Week', 'Total Orders', 'Total ShippedQty','Start Date', 'End Date',]
    # Formatear fechas como "DD-MMM" (ej: "22-Apr")
    df_weekly8['Start Date'] = df_weekly8['Start Date'].dt.strftime('%d-%b')
    df_weekly8['End Date'] = df_weekly8['End Date'].dt.strftime('%d-%b')
    rename_columns_weekly8 = {
        'Total ShippedQty': 'ASM Clubs',
        'Total Orders': 'ASM Orders',
    }
    df_weekly8 = df_weekly8.rename(columns=rename_columns_weekly8)
    df_weekly8 = df_weekly8 [['Week', 'Start Date', 'End Date', 'ASM Clubs', 'ASM Orders']]
    df_weekly8_data = [df_weekly8.columns.tolist()]
    df_weekly8_data += df_weekly8.values.tolist()
    df_weekly8_table = Table(df_weekly8_data)
    df_weekly8_table.setStyle(TableStyle(table_style_semana_actual))
    story.append(df_weekly8_table)

    story.append(PageBreak())

    #Misbuilds
    df_misbuild = df_semana_actual[df_semana_actual['Type'] == 'FRMISBUILD'] 
    misbuild_data = [df_misbuild.columns.tolist()]  # Encabezados
    misbuild_data += df_misbuild.values.tolist()    # Datos
    misbuild_table = Table(misbuild_data)
    misbuild_table.setStyle(TableStyle(table_style_semana_actual))
    story.append(misbuild_table)

    count_misbuilds = df4[df4['Type'] == 'FRMISBUILD']
    rename_columns_misbuilds = {
        'Claim Type (Description)': 'Description',
    }
    count_misbuilds = count_misbuilds.rename(columns=rename_columns_misbuilds)
    count_misbuilds = pd.crosstab(count_misbuilds['Description'], count_misbuilds['Historical Week'])
    count_misbuilds.loc['Total'] = count_misbuilds.sum(numeric_only=True)
    count_misbuilds_data =[['Count of Misbuilds']]
    count_misbuilds_data += [['Description'] + count_misbuilds.columns.tolist()]
    for idx, row in count_misbuilds.iterrows():
        count_misbuilds_data.append([idx] + row.astype(int).astype(str).tolist())
    #Tabla
    num_filas_cm = len(count_misbuilds_data)
    row_heights_cm = [13] * num_filas_cm
    count_misbuilds_table = Table(count_misbuilds_data, colWidths=[100, 60, 60, 60, 60], repeatRows=1, rowHeights=row_heights_cm)
    count_misbuilds_table.setStyle(TableStyle(table_style))

    #Avg Details
    week_cols_cm = [col for col in count_misbuilds.columns if col.startswith('Week')]
    count_misbuilds['TOTAL'] = count_misbuilds[week_cols_cm].sum(axis=1)
    non_null_weeks_countm = count_misbuilds[week_cols_cm].notnull().sum(axis=1)
    count_misbuilds['AVG'] = (count_misbuilds['TOTAL'] / non_null_weeks_countm).round(0).astype(int)
    avg_cm = count_misbuilds[['AVG','TOTAL']].copy()
    #Tabla 4 weeks
    avg_cm_data = [['Last 4 Weeks']]
    avg_cm_data += [list(avg_cm.columns)]
    avg_cm_data += avg_cm.values.tolist()
    num_filas_avg_cm = len(avg_cm_data)
    row_heights_avg_cm = [13] * num_filas_avg_cm
    avg_cm_table = Table(avg_cm_data, colWidths=[60,60], rowHeights=row_heights_avg_cm)
    avg_cm_table.setStyle(TableStyle(table_style_weeks))

    story.append(Spacer(width=0, height=1.5*cm))

    joined_cm = Table([[count_misbuilds_table, avg_cm_table]])
    story.append(joined_cm)
    story.append(Spacer(width=0, height=1.5*cm))
    # 1. Preparar los datos base (igual que antes)
    prod_data_cm = [
        ["Production data"],
        ["", *df_weekly['Week'].tolist()], 
        ["Start Date"] + df_weekly['Start Date'].tolist(),
        ["End Date"] + df_weekly['End Date'].tolist(),
        ["ASM Clubs"] + [f"{x:,.0f}" for x in df_weekly['Total ShippedQty']],
        ["Orders"] + [f"{x:,.0f}" for x in df_weekly['Total Orders']]
    ]

    # 2. Calcular las nuevas filas
    misbuilds_per_orders = []
    complemento_100 = []  # Nueva lista para 100% - Misbuilds/Orders

    for week in df_weekly['Week']:
        # Obtener el valor numérico REAL de misbuilds (3 métodos seguros)
        try:
            # Método 1: Si count_misbuilds es un DataFrame
            if isinstance(count_misbuilds, pd.DataFrame):
                misbuilds = count_misbuilds.loc[week].iloc[0] if week in count_misbuilds.index else 0
            # Método 2: Si es una Serie
            elif isinstance(count_misbuilds, pd.Series):
                misbuilds = count_misbuilds[week] if week in count_misbuilds.index else 0
            # Método 3: Si es un diccionario u otro tipo
            else:
                misbuilds = count_misbuilds.get(week, 0)
        except (KeyError, IndexError):
            misbuilds = 0
        
        # Convertir a float explícitamente (seguro)
        misbuilds = float(misbuilds)
        
        # Obtener orders (asegurarse que es un valor numérico)
        orders = float(df_weekly[df_weekly['Week'] == week]['Total Orders'].iloc[0])
        
        # Calcular ratios con protección completa
        ratio_orders = misbuilds / orders if orders != 0 else 0
        ratio_complemento = (1 - ratio_orders) if orders != 0 else 1
        
        # Formatear (Misbuilds/Orders como decimal, complemento como porcentaje)
        misbuilds_per_orders.append(f"{ratio_orders: .2%}")
        complemento_100.append(f"{ratio_complemento:.2%}")  # Formato de porcentaje

    # 3. Agregar las filas a la tabla
    prod_data_cm.append(["Misbuild of Orders"] + misbuilds_per_orders)
    prod_data_cm.append(["Build Quality"] + complemento_100)  # Nueva fila modificada

    # 4. Crear la tabla con ReportLab
    prod_cm_table = Table(
        prod_data_cm,
        colWidths=[100] + [60] * len(df_weekly),  # Ajusta según necesidad
    )
    prod_cm_table.setStyle(prod_style)
    prod_cm_joined = Table([[prod_cm_table, '']])
    story.append(prod_cm_joined)

    story.append(PageBreak())

    # HISTORICO MISBUILDS
    count_misbuilds8 = df8[df8['Type'] == 'FRMISBUILD']
    rename_columns_misbuilds8 = {
        'Claim Type (Description)': 'Description',
    }
    count_misbuilds8 = count_misbuilds8.rename(columns=rename_columns_misbuilds8)
    count_misbuilds8 = pd.crosstab(count_misbuilds8['Description'], count_misbuilds8['Historical Week'])
    count_misbuilds8.loc['Total'] = count_misbuilds8.sum(numeric_only=True)
    #count_misbuilds_data8 =[['Count of Misbuilds']]
    count_misbuilds_data8 = [['Description'] + count_misbuilds8.columns.tolist()]
    for idx, row in count_misbuilds8.iterrows():
        count_misbuilds_data8.append([idx] + row.astype(int).astype(str).tolist())

    # 3. Crear tabla de historial
    num_filas_misbuilds8 = len(count_misbuilds_data8)
    row_heights_mis8 = [13] * num_filas_misbuilds8
    count_misbuilds_table8 = Table(count_misbuilds_data8, colWidths=[120, 58, 58, 58, 58, 58], repeatRows=1, rowHeights=row_heights_mis8)
    count_misbuilds_table8.setStyle(TableStyle(table_style_graphic))
    # 4. Calcular los datos de resumen
    mb_last_8_weeks = count_misbuilds8.iloc[:, -8:].sum(axis=1) #if len(warranty_hist8.columns) <= 8 else pd.Series(0, index=warranty_hist8.index)
    mb_weeks_5_to_8 = count_misbuilds8.iloc[:, -8:-4].sum(axis=1) #if len(warranty_hist8.columns) <= 8 else pd.Series(0, index=warranty_hist8.index)
    mb_last_4_weeks = count_misbuilds8.iloc[:, -4:].sum(axis=1) #if len(warranty_hist8.columns) >= 4 else pd.Series(0, index=warranty_hist8.index)

    # 5. Preparar datos para la tabla de resumen
    summary_data_mis8 = [['Last 4 Weeks', 'Weeks 5-8', 'Dif','Total']]
    for idx in count_misbuilds8.index:
        summary_data_mis8.append([
            str(int(mb_last_4_weeks[idx])),
            str(int(mb_weeks_5_to_8[idx])),
            str(int(mb_last_4_weeks[idx])-int(mb_weeks_5_to_8[idx])),
            str(int(mb_last_8_weeks[idx]))        
        ])

    # 6. Crear tabla de resumen
    summary_table_mis8 = Table(summary_data_mis8, colWidths=[58, 58, 58, 58], repeatRows=1, rowHeights=row_heights_mis8)
    summary_table_mis8.setStyle(TableStyle(table_style_graphic2))
    graphic_joined_mis8 = Table([[count_misbuilds_table8, summary_table_mis8]])
    story.append(graphic_joined_mis8)

    #GRAFICA
    # Gráfico de líneas
    misbuilds_counts = df[df['Type'] == 'FRMISBUILD'].groupby('Historical Week').size()
    misbuilds_counts = misbuilds_counts.reindex(df_weekly['Week'], fill_value=0)
    # 2. Calcular el promedio móvil de 4 semanas
    misbuilds_4wk_avg = misbuilds_counts.rolling(window=4, min_periods=1).mean()
    # Crear figura y eje principal (misbuilds)
    fig, ax1 = plt.subplots(figsize=(12, 6))

    # Gráfico de Misbuilds (eje izquierdo - rojo)
    ax1.plot(misbuilds_counts.index, misbuilds_counts.values, 
            label='Total Misbuilds', 
            color='red', 
            marker='s', 
            linestyle='--', 
            linewidth=2)

    # Gráfico de Promedio Móvil (eje izquierdo - verde)
    ax1.plot(misbuilds_4wk_avg.index, misbuilds_4wk_avg.values, 
            label='4 Week Avg', 
            color='green', 
            marker='^', 
            linestyle='-.', 
            linewidth=2)
    ax1.tick_params(axis='y')
    ax1.grid(True, linestyle='--', alpha=0.3)

    # Crear eje secundario (derecho - azul)
    ax2 = ax1.twinx()

    # Gráfico de Total de ordenes (eje derecho - azul)
    ax2.plot(df_weekly['Week'], df_weekly['Total Orders'], 
            label='Total Orders', 
            color='blue', 
            marker='o', 
            linestyle='-', 
            linewidth=2)
    ax2.tick_params(axis='y')

    # Título y leyenda unificada
    plt.title('Misbuilds and Orders over Time', fontsize=14, pad=20)

    # Combinar leyendas
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=10, loc='upper left')

    # Ajustar formato
    plt.xticks(rotation=45)
    fig.tight_layout()

    # Guardar imagen temporal
    plt.savefig("graphic2.png", dpi=300)
    plt.close()

    # Insertar la imagen (gráfica)
    story.append(Spacer(width=0, height=1.5*cm))
    story.append(Image("graphic2.png", width=750, height=300)) 
    # Guardar
    doc.build(story, onFirstPage=draw_cover)

    return doc

def main():
    st.title("📊 Generador de Reporte de Defectos y Garantías")
    
    # Subida de archivos
    defect_file = st.file_uploader("Archivo de Defectos (Excel)", type=['xlsx', 'xls'])
    production_file = st.file_uploader("Archivo de Producción (CSV)", type=["csv"])
    
    if defect_file and production_file:
        with st.spinner('Generando reporte...'):
            try:
                # Procesar archivos y obtener el PDF en bytes
                pdf_bytes = procesar_archivos(defect_file, production_file)
                
                # Mostrar botón de descarga
                st.download_button(
                    label="⬇️ Descargar Reporte",
                    data=pdf_bytes,
                    file_name="reporte_defectos.pdf",
                    mime="application/pdf"
                )
                
                st.success("¡Reporte generado con éxito!")
                
            except Exception as e:
                st.error(f"Error: {str(e)}")

def procesar_archivos(defect_file, production_file):
    # Crear buffer para el PDF
    pdf_buffer = io.BytesIO()
    
    # Configurar el documento
    doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
    
    # Aquí tu lógica de procesamiento...
    story = []
    # story.append(...) - Agrega todos tus elementos al story
    
    # Construir el PDF
    doc.build(story)
    
    # Obtener los bytes del PDF
    pdf_buffer.seek(0)
    return pdf_buffer

if __name__ == "__main__":
    main()