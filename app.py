import streamlit as st
import pandas as pd
import io

st.set_page_config(layout="wide")
st.title("📊 Контроль стандартов (жёсткая модель пациента)")

uploaded_file = st.file_uploader("Загрузите Excel", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    # -----------------------------
    # ОЧИСТКА
    # -----------------------------
    df = df.dropna(subset=["PATIENTS_ID", "Код услуги"])

    df = df[
        ~df["Название услуги"]
        .str.lower()
        .str.contains("прием|консультац", na=False)
    ]

    # -----------------------------
    # 1. PATIENT × SERVICE (БАЗА)
    # -----------------------------
    ps = (
        df.groupby(["PATIENTS_ID","Код услуги","Название услуги"], as_index=False)
        .agg(total=("Назначено","sum"))
    )

    ps["done"] = (ps["total"] > 0).astype(int)

    # -----------------------------
    # 2. УРОВЕНЬ УСЛУГ (ГЛОБАЛЬНЫЙ СЛОЙ)
    # -----------------------------
    svc_base = (
        ps.groupby(["Код услуги","Название услуги"], as_index=False)
        .agg(
            patients_need=("PATIENTS_ID","nunique"),
            total_assignments=("total","sum")
        )
    )
    
    received = ps[ps["total"] > 0].groupby(["Код услуги","Название услуги"]).agg(
        patients_received=("PATIENTS_ID", "nunique")
    ).reset_index()
    
    svc = svc_base.merge(received, on=["Код услуги","Название услуги"], how="left")
    svc["patients_received"] = svc["patients_received"].fillna(0)
    
    svc["overuse_ratio"] = svc["total_assignments"] / svc["patients_received"].replace(0, 1)
    svc["overuse_flag"] = (svc["overuse_ratio"] > 1).astype(int)
    svc["coverage_%"] = (svc["patients_received"] / svc["patients_need"] * 100).round(1)
    svc["patients_done"] = svc["patients_received"]
    
    overuse_services_codes = svc[svc["overuse_flag"] == 1]["Код услуги"].tolist()

    # -----------------------------
    # 3. PS ДЛЯ БИНАРНЫХ РАСЧЕТОВ
    # -----------------------------
    ps_binary = ps.copy()
    ps_binary["global_done"] = ps_binary["done"]
    ps_binary["service_overuse"] = ps_binary["Код услуги"].isin(overuse_services_codes).astype(int)
    ps_binary["patient_overuse"] = (ps_binary["service_overuse"] == 1) & (ps_binary["total"] > 1)
    ps_binary["patient_overuse"] = ps_binary["patient_overuse"].astype(int)

    # -----------------------------
    # KPI
    # -----------------------------
    total_pairs = len(ps)
    not_assigned_pairs = (ps["total"] == 0).sum()
    
    total_services = len(svc)
    overuse_services = svc["overuse_flag"].sum()
    
    patient_stats = (
        ps_binary.groupby("PATIENTS_ID")
        .agg(
            has_done=("global_done","max"),
            has_overuse=("patient_overuse","max")
        )
    )

    patients_with_services = patient_stats["has_done"].sum()
    overuse_patients = patient_stats["has_overuse"].sum()

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("❌ % не назначенных услуг", 
                f"{not_assigned_pairs/total_pairs*100:.1f}%")
    
    col2.metric("⚠️ % избыточных услуг", 
                f"{overuse_services/total_services*100:.1f}%")
    
    col3.metric("👤 пациенты с услугами", 
                int(patients_with_services))
    
    col4.metric("⚠️ % избыточных пациентов",
                f"{overuse_patients / max(patients_with_services,1) * 100:.1f}%")

    st.markdown("---")

    # -----------------------------
    # ТОП ИЗБЫТОЧНЫХ УСЛУГ
    # -----------------------------
    st.subheader("🔥 Избыточные услуги (среднее назначений на получившего > 1)")
    
    overuse_calc = ps[ps["total"] > 0].groupby(["Код услуги", "Название услуги"]).agg(
        patients_received=("PATIENTS_ID", "nunique"),
        total_assignments=("total", "sum"),
        avg_per_patient=("total", "mean")
    ).reset_index()
    
    need_info = ps.groupby(["Код услуги", "Название услуги"]).agg(
        patients_need=("PATIENTS_ID", "nunique")
    ).reset_index()
    
    overuse_calc = overuse_calc.merge(need_info, on=["Код услуги", "Название услуги"], how="left")
    overuse_calc["overuse_ratio"] = overuse_calc["avg_per_patient"].round(2)
    overuse_calc["overuse_flag"] = (overuse_calc["overuse_ratio"] > 1).astype(int)
    overuse_calc["отклонение"] = (overuse_calc["overuse_ratio"] - 1).round(2)
    overuse_calc["coverage_%"] = (overuse_calc["patients_received"] / overuse_calc["patients_need"] * 100).round(1)
    
    overuse_only = overuse_calc[overuse_calc["overuse_flag"] == 1].copy()
    
    if len(overuse_only) > 0:
        st.dataframe(
            overuse_only.sort_values("overuse_ratio", ascending=False)[
                ["Название услуги", "Код услуги", "patients_need", "patients_received", 
                 "total_assignments", "overuse_ratio", "отклонение", "coverage_%"]
            ],
            use_container_width=True,
            height=400
        )
        st.caption(f"📊 Всего избыточных услуг: {len(overuse_only)}")
    else:
        st.info("✅ Нет избыточных услуг.")
    
    st.markdown("---")

    # -----------------------------
    # ФИЛЬТРЫ
    # -----------------------------
    st.subheader("⚙️ Фильтры")

    all_departments = sorted(df["Отделение"].dropna().unique())

    selected_departments = st.multiselect(
        "Отделения",
        options=all_departments,
        default=all_departments
    )

    df_f = df[df["Отделение"].isin(selected_departments)]

    st.markdown("---")

    # -----------------------------
    # 4. ПРОЕКЦИЯ В ОТДЕЛЕНИЯ (БИНАРНАЯ ЛОГИКА)
    # -----------------------------
    
    # Карта: есть ли услуга у пациента (глобально)
    has_service_map = ps[["PATIENTS_ID", "Код услуги", "done"]].copy()
    has_service_map = has_service_map.rename(columns={"done": "has_service"})
    
    # Данные по отделениям
    dept_data = df_f.groupby(
        ["PATIENTS_ID", "Код услуги", "Название услуги", "Отделение"], 
        as_index=False
    ).agg(
        total_in_dept=("Назначено", "sum")
    )
    
    # Добавляем глобальную метку
    dept_data = dept_data.merge(has_service_map, on=["PATIENTS_ID", "Код услуги"], how="left")
    dept_data["has_service"] = dept_data["has_service"].fillna(0).astype(int)
    
    # Агрегация по отделениям
    dept = (
        dept_data.groupby(["Отделение", "Код услуги", "Название услуги"], as_index=False)
        .agg(
            patients=("PATIENTS_ID", "nunique"),
            patients_done=("has_service", "sum"),
            total=("total_in_dept", "sum")
        )
    )
    
    dept["% выполнения"] = (dept["patients_done"] / dept["patients"] * 100).round(1)
    dept["% невыполнения"] = (100 - dept["% выполнения"]).round(1)
    dept["avg_assignments"] = (dept["total"] / dept["patients"]).round(2)
    dept["overuse_flag"] = (dept["total"] > dept["patients"]).astype(int)
    
    st.subheader("🏥 Отделения (детализация по услугам)")
    st.dataframe(
        dept.sort_values("% невыполнения", ascending=False),
        use_container_width=True,
        height=400
    )
    
    # Сводка по отделениям
    dept_summary = (
        dept.groupby("Отделение")
        .agg(
            total_services=("Код услуги", "nunique"),
            total_patients=("patients", "sum"),
            total_patients_done=("patients_done", "sum")
        )
        .reset_index()
    )
    dept_summary["% выполнения"] = (dept_summary["total_patients_done"] / dept_summary["total_patients"] * 100).round(1)
    dept_summary = dept_summary.sort_values("% выполнения", ascending=False)
    
    st.subheader("📊 Сводка по отделениям")
    st.dataframe(dept_summary, use_container_width=True)
    
    st.markdown("---")
    
    # -----------------------------
    # 5. ПРОЕКЦИЯ В ПЛАНЫ (БИНАРНАЯ ЛОГИКА)
    # -----------------------------
    
    # Данные по планам
    plan_data = df_f.groupby(
        ["PATIENTS_ID", "Код услуги", "Название услуги", "Название плана лечения"], 
        as_index=False
    ).agg(
        total_in_plan=("Назначено", "sum")
    )
    
    # Добавляем глобальную метку
    plan_data = plan_data.merge(has_service_map, on=["PATIENTS_ID", "Код услуги"], how="left")
    plan_data["has_service"] = plan_data["has_service"].fillna(0).astype(int)
    
    # Агрегация по планам
    plan = (
        plan_data.groupby(["Название плана лечения", "Код услуги", "Название услуги"], as_index=False)
        .agg(
            patients=("PATIENTS_ID", "nunique"),
            patients_done=("has_service", "sum"),
            total=("total_in_plan", "sum")
        )
    )
    
    plan["% выполнения"] = (plan["patients_done"] / plan["patients"] * 100).round(1)
    plan["% невыполнения"] = (100 - plan["% выполнения"]).round(1)
    plan["avg_assignments"] = (plan["total"] / plan["patients"]).round(2)
    plan["overuse_flag"] = (plan["total"] > plan["patients"]).astype(int)
    
    st.subheader("📂 Планы лечения (детализация по услугам)")
    st.dataframe(
        plan.sort_values("% невыполнения", ascending=False),
        use_container_width=True,
        height=400
    )
    
    # Сводка по планам
    plan_summary = (
        plan.groupby("Название плана лечения")
        .agg(
            total_services=("Код услуги", "nunique"),
            total_patients=("patients", "sum"),
            total_patients_done=("patients_done", "sum")
        )
        .reset_index()
    )
    plan_summary["% выполнения"] = (plan_summary["total_patients_done"] / plan_summary["total_patients"] * 100).round(1)
    plan_summary = plan_summary.sort_values("% выполнения", ascending=False)
    
    st.subheader("📊 Сводка по планам лечения")
    st.dataframe(plan_summary, use_container_width=True)
    
    st.markdown("---")

    # -----------------------------
    # 6. УСЛУГИ (ГЛОБАЛЬНО)
    # -----------------------------
    st.subheader("🧪 Услуги (общий слой)")
    
    display_svc = svc.copy()
    display_svc["overuse_ratio"] = display_svc["overuse_ratio"].round(2)
    display_svc["coverage_%"] = display_svc["coverage_%"].round(1)
    
    st.dataframe(
        display_svc.sort_values("overuse_ratio", ascending=False),
        use_container_width=True,
        height=500
    )

    # -----------------------------
    # 7. ЭКСПОРТ
    # -----------------------------
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        svc.to_excel(writer, sheet_name="services", index=False)
        dept.to_excel(writer, sheet_name="departments", index=False)
        dept_summary.to_excel(writer, sheet_name="departments_summary", index=False)
        plan.to_excel(writer, sheet_name="plans", index=False)
        plan_summary.to_excel(writer, sheet_name="plans_summary", index=False)
        ps_binary.to_excel(writer, sheet_name="patient_service", index=False)

    output.seek(0)

    st.download_button(
        label="📥 Скачать Excel",
        data=output,
        file_name="analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
