"""
Car Affordability Calculator — Free Web Tool by ClearMetric
https://clearmetric.gumroad.com

Product T8. Helps people figure out how much car they can actually afford.
"""

import streamlit as st
import plotly.graph_objects as go
import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Car Affordability Calculator — ClearMetric",
    page_icon="🚗",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Custom CSS (navy theme)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main .block-container { padding-top: 2rem; max-width: 1200px; }
    .stMetric { background: #f8f9fa; border-radius: 8px; padding: 12px; border-left: 4px solid #2C3E50; }
    h1 { color: #2C3E50; }
    h2, h3 { color: #1C2833; }
    .rule-pass { color: #27ae60; font-weight: bold; }
    .rule-fail { color: #e74c3c; font-weight: bold; }
    .cta-box {
        background: linear-gradient(135deg, #2C3E50 0%, #1C2833 100%);
        color: white; padding: 24px; border-radius: 12px; text-align: center;
        margin: 20px 0;
    }
    .cta-box a { color: #D5D8DC; text-decoration: none; font-weight: bold; font-size: 1.1rem; }
    div[data-testid="stSidebar"] { background: #f8f9fa; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Credit score → suggested rate
# ---------------------------------------------------------------------------
CREDIT_RATES = {
    "Excellent (720+)": 5.5,
    "Good (670-719)": 6.5,
    "Fair (580-669)": 9.0,
    "Poor (<580)": 12.0,
}

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown("# 🚗 Car Affordability Calculator")
st.markdown("**How much car can you actually afford?** Based on the 20/4/10 rule and total cost of ownership.")
st.markdown("---")

# ---------------------------------------------------------------------------
# Sidebar — User inputs
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## Your Finances")
    st.button("🔄 Update Results", use_container_width=True)

    st.markdown("### Income & Debt")
    monthly_income = st.number_input(
        "Monthly Take-Home Income ($)",
        value=5_000,
        min_value=0,
        step=500,
    )
    monthly_debt = st.number_input(
        "Current Monthly Debt Payments ($)",
        value=500,
        min_value=0,
        step=50,
    )
    max_car_pct = st.slider(
        "Max % of Income for Car (20/4/10 = 10%)",
        10,
        25,
        15,
        1,
    ) / 100

    st.markdown("### Down Payment & Trade-In")
    down_payment = st.number_input(
        "Down Payment Available ($)",
        value=5_000,
        min_value=0,
        step=500,
    )
    trade_in = st.number_input(
        "Trade-In Value ($)",
        value=0,
        min_value=0,
        step=500,
    )

    st.markdown("### Loan Terms")
    loan_term = st.selectbox(
        "Loan Term",
        [36, 48, 60, 72],
        index=2,
        format_func=lambda x: f"{x} months",
    )
    credit_score = st.selectbox(
        "Credit Score Range",
        list(CREDIT_RATES.keys()),
        index=1,
    )
    interest_rate = st.number_input(
        "Interest Rate (%)",
        value=float(CREDIT_RATES[credit_score]),
        min_value=0.0,
        max_value=25.0,
        step=0.25,
        format="%.2f",
    ) / 100

    st.markdown("### Other Costs")
    include_insurance = st.checkbox("Include Insurance Estimate?", value=True)
    insurance_monthly = 150
    if include_insurance:
        insurance_monthly = st.number_input(
            "Monthly Insurance Estimate ($)",
            value=150,
            min_value=0,
            step=25,
        )
    annual_maintenance = st.number_input(
        "Annual Maintenance Budget ($)",
        value=1_200,
        min_value=0,
        step=100,
    )
    annual_fuel = st.number_input(
        "Annual Fuel Cost ($)",
        value=2_400,
        min_value=0,
        step=200,
    )
    sales_tax_rate = st.number_input(
        "Sales Tax Rate (%)",
        value=7.0,
        min_value=0.0,
        max_value=15.0,
        step=0.5,
        format="%.1f",
    ) / 100

# ---------------------------------------------------------------------------
# Core calculations
# ---------------------------------------------------------------------------
# Max affordable monthly payment (P&I only)
max_monthly_payment = max(0, monthly_income * max_car_pct - monthly_debt)

# Max loan amount (PMT formula backwards: PV = PMT * ((1 - (1+r)^(-n)) / r))
n_months = loan_term
monthly_rate = interest_rate / 12
if monthly_rate > 0 and max_monthly_payment > 0:
    max_loan = max_monthly_payment * (
        (1 - (1 + monthly_rate) ** (-n_months)) / monthly_rate
    )
else:
    max_loan = max_monthly_payment * n_months  # 0% rate fallback

# Max car price: loan = (price * (1 + tax)) - down - trade_in
# price = (loan + down + trade_in) / (1 + tax)
max_car_price = (max_loan + down_payment + trade_in) / (1 + sales_tax_rate)

# Depreciation estimate (first year 20%, then 15%/year)
years = n_months / 12
depreciation_by_year = []
val = max_car_price
for y in range(int(years) + 1):
    depreciation_by_year.append({"Year": y, "Value": val})
    if y == 0:
        val = val * 0.80  # 20% first year
    else:
        val = val * 0.85  # 15% thereafter
depreciation_df = pd.DataFrame(depreciation_by_year)

# Total cost of ownership over loan term
total_payments = max_monthly_payment * n_months
total_insurance = insurance_monthly * n_months if include_insurance else 0
total_maintenance = annual_maintenance * (n_months / 12)
total_fuel = annual_fuel * (n_months / 12)
end_value = depreciation_df.iloc[-1]["Value"] if len(depreciation_df) > 0 else 0
total_depreciation = max_car_price - end_value
total_tco = total_payments + total_insurance + total_maintenance + total_fuel

# New vs Used: same budget
# New: max_car_price
# 3-yr-old: same OTD budget. A 3-yr-old car = 0.8 * 0.85^2 = 0.578 of original.
# So we can afford a used car worth max_car_price. That car was new at max_car_price / 0.578
used_multiplier = 1 / (0.80 * 0.85 * 0.85)  # ~1.73
equivalent_new_price_used = max_car_price * used_multiplier

# 20/4/10 rule check
# 20% down, 4-year loan, 10% of income on total car costs
rule_20_down = down_payment >= (max_car_price * 0.20)
rule_4_year = loan_term <= 48
total_car_costs_monthly = (
    max_monthly_payment
    + (insurance_monthly if include_insurance else 0)
    + annual_maintenance / 12
    + annual_fuel / 12
)
rule_10_pct = total_car_costs_monthly <= (monthly_income * 0.10)
rule_20410_pass = rule_20_down and rule_4_year and rule_10_pct

# True monthly cost
true_monthly_cost = (
    max_monthly_payment
    + (insurance_monthly if include_insurance else 0)
    + annual_maintenance / 12
    + annual_fuel / 12
)

# ---------------------------------------------------------------------------
# Display — Key metrics
# ---------------------------------------------------------------------------
st.markdown("## Key Results")

m1, m2, m3, m4 = st.columns(4)
m1.metric(
    "Max Car Price",
    f"${max_car_price:,.0f}",
    help="Maximum vehicle price you can afford (before tax)",
)
m2.metric(
    "Monthly Payment",
    f"${max_monthly_payment:,.0f}",
    help="Estimated loan payment (P&I only)",
)
m3.metric(
    "Total Cost (Loan Term)",
    f"${total_tco:,.0f}",
    help="Payments + insurance + maintenance + fuel",
)
rule_label = "✓ Pass" if rule_20410_pass else "✗ Fail"
m4.metric(
    "20/4/10 Rule",
    rule_label,
    help="20% down, 4-yr loan, 10% income on car costs",
)

st.markdown("---")

# ---------------------------------------------------------------------------
# Affordability gauge
# ---------------------------------------------------------------------------
st.markdown("## Affordability Gauge")
gauge_max = max(60_000, int(max_car_price * 1.2))
fig_gauge = go.Figure(
    go.Indicator(
        mode="gauge+number",
        value=max_car_price,
        number={"prefix": "$", "suffix": "", "font": {"size": 28}},
        title={"text": "Max Affordable Car Price"},
        gauge={
            "axis": {"range": [0, gauge_max], "tickprefix": "$"},
            "bar": {"color": "#2C3E50"},
            "steps": [
                {"range": [0, gauge_max * 0.25], "color": "#E8F6F3"},
                {"range": [gauge_max * 0.25, gauge_max * 0.5], "color": "#D5D8DC"},
                {"range": [gauge_max * 0.5, gauge_max * 0.75], "color": "#AEB6BF"},
                {"range": [gauge_max * 0.75, gauge_max], "color": "#5D6D7E"},
            ],
            "threshold": {
                "line": {"color": "#1C2833", "width": 4},
                "thickness": 0.75,
                "value": max_car_price,
            },
        },
    )
)
fig_gauge.update_layout(height=280, margin=dict(l=20, r=20, t=50, b=20))
st.plotly_chart(fig_gauge, use_container_width=True)

# ---------------------------------------------------------------------------
# New vs Used comparison
# ---------------------------------------------------------------------------
st.markdown("## New vs Used: Same Budget")
col1, col2 = st.columns(2)
with col1:
    st.markdown("### New Car")
    st.metric("Max Price", f"${max_car_price:,.0f}")
    st.caption("Brand new vehicle at your max budget")
with col2:
    st.markdown("### 3-Year-Old Used")
    st.metric(
        "Equivalent New Price",
        f"${equivalent_new_price_used:,.0f}",
    )
    st.caption(
        "A 3-year-old car worth the same OTD — originally cost ~"
        f"${equivalent_new_price_used:,.0f} when new"
    )

# ---------------------------------------------------------------------------
# Total cost breakdown pie
# ---------------------------------------------------------------------------
st.markdown("## Total Cost Breakdown (Loan Term)")
pie_labels = ["Loan Payments", "Insurance", "Maintenance", "Fuel"]
pie_values = [
    total_payments,
    total_insurance,
    total_maintenance,
    total_fuel,
]
# Filter out zeros
pie_data = [(l, v) for l, v in zip(pie_labels, pie_values) if v > 0]
if pie_data:
    fig_pie = go.Figure(
        data=[
            go.Pie(
                labels=[x[0] for x in pie_data],
                values=[x[1] for x in pie_data],
                hole=0.4,
                marker_colors=["#2C3E50", "#5D6D7E", "#85929E", "#AEB6BF"],
            )
        ]
    )
    fig_pie.update_layout(
        height=400,
        showlegend=True,
        legend=dict(orientation="h", y=1.02),
        margin=dict(t=40, b=40),
    )
    st.plotly_chart(fig_pie, use_container_width=True)

# ---------------------------------------------------------------------------
# Depreciation curve
# ---------------------------------------------------------------------------
st.markdown("## Estimated Depreciation (20% yr1, 15%/yr after)")
fig_dep = go.Figure()
fig_dep.add_trace(
    go.Scatter(
        x=depreciation_df["Year"],
        y=depreciation_df["Value"],
        mode="lines+markers",
        line=dict(color="#2C3E50", width=3),
        marker=dict(size=8),
    )
)
fig_dep.update_layout(
    yaxis_title="Estimated Value ($)",
    yaxis_tickformat="$,.0f",
    xaxis_title="Years Owned",
    height=400,
    template="plotly_white",
    margin=dict(t=40, b=40),
)
st.plotly_chart(fig_dep, use_container_width=True)

# ---------------------------------------------------------------------------
# 20/4/10 rule details
# ---------------------------------------------------------------------------
st.markdown("## 20/4/10 Rule Check")
if rule_20410_pass:
    st.success(
        "**You pass the 20/4/10 rule.** 20% down, 4-year loan or less, "
        "and total car costs ≤10% of income."
    )
else:
    reasons = []
    if not rule_20_down:
        needed = max_car_price * 0.20
        reasons.append(f"20% down: need ${needed:,.0f}, you have ${down_payment:,.0f}")
    if not rule_4_year:
        reasons.append(f"4-year loan: you chose {loan_term} months")
    if not rule_10_pct:
        pct = (total_car_costs_monthly / monthly_income) * 100
        reasons.append(
            f"10% of income: car costs are {pct:.1f}% of income "
            f"(${total_car_costs_monthly:,.0f}/mo)"
        )
    st.warning(
        "**You don't pass the 20/4/10 rule.** "
        + "; ".join(reasons)
    )

st.markdown("**True Monthly Cost:** " + f"${true_monthly_cost:,.0f}/mo (payment + insurance + maintenance + fuel)")

# ---------------------------------------------------------------------------
# CTA — Paid Excel
# ---------------------------------------------------------------------------
st.markdown("---")
st.markdown("""
<div class="cta-box">
    <h3 style="color: white; margin: 0 0 8px 0;">Want the Full Excel Spreadsheet?</h3>
    <p style="margin: 0 0 16px 0;">
        Get the <strong>ClearMetric Car Affordability Calculator</strong> — $8.99<br>
        ✓ All inputs in one place with editable cells<br>
        ✓ Compare 3 cars side by side (price, loan, insurance, fuel, maintenance, depreciation)<br>
        ✓ 20/4/10 rule checker + How To Use guide<br>
    </p>
    <a href="https://clearmetric.gumroad.com/l/car-affordability" target="_blank">
        Get It on Gumroad — $8.99 →
    </a>
</div>
""", unsafe_allow_html=True)

# Cross-sell Budget Planner
st.markdown("### More from ClearMetric")
cx1, cx2, cx3 = st.columns(3)
with cx1:
    st.markdown("""
    **📊 Budget Planner** — $13.99
    Track income, expenses, savings with the 50/30/20 framework.
    [Get it →](https://clearmetric.gumroad.com/l/budget-planner)
    """)
with cx2:
    st.markdown("""
    **🏠 Rent vs Buy** — $12.99
    Should you rent or buy? Compare true financial impact.
    [Get it →](https://clearmetric.gumroad.com/l/rent-vs-buy)
    """)
with cx3:
    st.markdown("""
    **🔥 FIRE Calculator** — $14.99
    Find your FIRE number, scenario comparison.
    [Get it →](https://clearmetric.gumroad.com/l/fire-calculator)
    """)

# Footer
st.markdown("---")
st.caption(
    "© 2026 ClearMetric | [clearmetric.gumroad.com](https://clearmetric.gumroad.com) | "
    "This tool is for educational purposes only. Not financial advice."
)
