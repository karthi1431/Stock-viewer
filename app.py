from flask import Flask, render_template, request, send_file
from nselib import capital_market
import pandas as pd
import plotly.express as px
import plotly.io as pio
import yfinance as yf
import os
import math

app = Flask(__name__)

PAGE_SIZE = 30


def price_category(price):
    if price < 100:
        return "Low"
    elif price < 500:
        return "Mid"
    else:
        return "High"


def trend_label(p):
    if p > 2:
        return "Strong Gainer"
    elif p > 0:
        return "Mild Gainer"
    elif p > -2:
        return "Flat"
    else:
        return "Loser"


@app.route("/", methods=["GET", "POST"])
def index():
    trade_date = ""
    stock_symbol = ""
    page = 1
    error = None
    table_html = None
    stock_table_html = None
    stock_msg = None
    fig1_html = fig2_html = fig3_html = fig4_html = None

    df_display = None
    total_pages = 1

    if request.method == "POST":
        trade_date = request.form.get("trade_date", "").strip()
        stock_symbol = request.form.get("stock_symbol", "").strip().upper()
        page = int(request.form.get("page", 1))

        try:
            data = capital_market.bhav_copy_equities(trade_date=trade_date)
            df = data.copy() if isinstance(data, pd.DataFrame) else pd.DataFrame(data)

            if df.empty:
                error = "No data for this date."
                return render_template("index.html", error=error)

            # Auto column detection
            def pick(cols):
                for c in cols:
                    if c in df.columns:
                        return c
                return None

            symbol_col = pick(["SYMBOL", "SctySym", "FinInstrmNm"])
            open_col = pick(["OPEN_PRICE", "OpnPric"])
            close_col = pick(["CLOSE_PRICE", "ClsPric"])
            val_col = pick(["TTL_TRD_VAL", "TradVal", "TotalTradedValue"])
            sector_col = pick(["IndNm", "Sector"])

            display_cols = [c for c in [symbol_col, open_col, close_col, val_col, sector_col] if c]
            df_display = df[display_cols].copy()

            # Categorization
            df_display[open_col] = pd.to_numeric(df_display[open_col], errors="coerce")
            df_display[close_col] = pd.to_numeric(df_display[close_col], errors="coerce")

            df_display["PRICE_CATEGORY"] = df_display[close_col].apply(price_category)
            df_display["PCT_CHANGE"] = ((df_display[close_col] - df_display[open_col]) / df_display[open_col]) * 100
            df_display["TREND"] = df_display["PCT_CHANGE"].apply(trend_label)

            # Pagination
            total_rows = len(df_display)
            total_pages = math.ceil(total_rows / PAGE_SIZE)
            start = (page - 1) * PAGE_SIZE
            end = start + PAGE_SIZE

            df_page = df_display.iloc[start:end]
            table_html = df_page.to_html(classes="table table-striped table-sm", index=False)

            # ✅ Download page to Excel
            df_page.to_excel("current_page.xlsx", index=False)

            # ✅ Stock Search
            if stock_symbol:
                df_stock = df_display[df_display[symbol_col].astype(str).str.upper() == stock_symbol]

                if not df_stock.empty:
                    stock_table_html = df_stock.to_html(classes="table table-bordered table-sm", index=False)
                    stock_msg = f"Found {len(df_stock)} row(s) for {stock_symbol}"

                    # ✅ Multi-day chart using Yahoo Finance
                    yf_symbol = stock_symbol + ".NS"
                    hist = yf.Ticker(yf_symbol).history(period="1mo")

                    fig4 = px.line(hist, x=hist.index, y="Close",
                                    title=f"{stock_symbol} – Last 1 Month Price")
                    fig4_html = pio.to_html(fig4, full_html=False)

                else:
                    stock_msg = "Stock not found."

            # ✅ Top 10 by traded value
            if val_col:
                top = df_display.sort_values(val_col, ascending=False).head(10)
                fig1 = px.bar(top, x=symbol_col, y=val_col,
                              title="Top 10 by Traded Value")
                fig1_html = pio.to_html(fig1, full_html=False)

            # ✅ Sector-wise chart
            if sector_col:
                sector_counts = df_display[sector_col].value_counts().reset_index()
                sector_counts.columns = ["Sector", "Count"]
                fig3 = px.pie(sector_counts, names="Sector", values="Count",
                              title="Sector Distribution")
                fig3_html = pio.to_html(fig3, full_html=False)

        except Exception as e:
            error = str(e)

    return render_template(
        "index.html",
        trade_date=trade_date,
        stock_symbol=stock_symbol,
        page=page,
        total_pages=total_pages,
        table_html=table_html,
        stock_table_html=stock_table_html,
        stock_msg=stock_msg,
        fig1=fig1_html,
        fig3=fig3_html,
        fig4=fig4_html,
        error=error
    )


@app.route("/download")
def download():
    return send_file("current_page.xlsx", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
