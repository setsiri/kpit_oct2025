# kpit_oct2025
จากที่พิจารณาแบบ ดูว่า ตัวแปร SP ตรงกับ row1: A1,J1,V1 อันไหน หรือไม่ตรงเลย
ให้อ่านทั้ง row1 ก่อนว่ามี cell ไหนที่ไม่ใช่ blank บ้างและอยู่ที่ ช่องอะไรบ้าง
สุมมติ ถ้าเจอ 3 ค่า แล้วอยู่ A1 Q1 Y1 ก็ให้พิจารณาแบบ ดูว่า ตัวแปร SP ตรงกับ row1:  A1, Q1, Y1 หรือไม่ตรงเลย เป็นต้น

และ

จากที่พิจารณาแบบ ดูว่า ถ้าตรง A1 ให้ไล่อ่าน row2: D2,E2,F2,G2 ว่า ตัวแปร SOP ตรงกับ อันไหนหรือไม่ตรงเลย
ให้อ่านทั้ง row2 ตั้งแต่ A2 ถึง ก่อนหน้า Q2 ก็คือ P2


ถ้าตัวอย่างคือ พบ  row1:  A1, Q1, Y1
ก็จะเป็น
ถ้าตรง A1 ให้ไล่ทำ row2:A2-P2
ถ้าตรง Q1 ให้ไล่ทำ row2: Q2-X2
ถ้าตรง Y1 ให้ไล่ทำ row2: Y2 ถึงจนกว่าจะเจอค่า blank ครั้งแรก

ก็คือเปลี่ยนจากแบบ ฟิกช่องแน่นอน ให้มีความยืดหยุ่น 


ปรับการประมวลผลของ sheet ให้ทำเฉพาะ sheet ที่ 5 - 22 ก็พอ
และให้เอาจาก result ที่ได้มา represent ผลลัพท์อีกต่อหนึ่ง
โดยแสดงในรูปแบบนี้

Results: Each pages
----------------------------------------------
5. L1_FI -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
6. FuSa_Charger -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = False
7. FuSa_NACS -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
8. FuSa_FTT -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
9. FuSa_14DCDC -> Yes_All_applicable_TC_Presented = True; No_Inapplicable_TC_Presented = True
10. FuSa_14DCDC_EFAN -> Yes_All_applicable_TC_Presented = True; No_Inapplicable_TC_Presented = True
11. FuSa_SEV -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
12. FuSa_FaultInjection -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
13. Function_Charger -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
14. Function_SCC -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
15. Function_14DCDC -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
16. LEV4ZEV -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
17. Thermal Management -> Yes_All_applicable_TC_Presented = True; No_Inapplicable_TC_Presented = False
18. HW_SOP1.6_Delta_TCs -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
19. HW_SOP2.0_Delta_TCs -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
20. HW_SOP3.0_Delta_TCs -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
21. HW_SOP3.0_Delta_TCs_Pratik -> Yes_All_applicable_TC_Presented = True; No_Inapplicable_TC_Presented = True
22. SW_Delta_TCs -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
----------------------------------------------
Result: Test Level
L1_FI -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
L2_FuSa -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = False
L3_QM -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
HW_Delta1.6 -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
HW_Delta2.0 -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True
HW_Delta3.0 -> Yes_All_applicable_TC_Presented = False; No_Inapplicable_TC_Presented = True

โดย 
L1_FI ดูจาก ผลของ sheet เบอร์ 5.
L2_FuSa ดูจาก ผลของ sheet เบอร์ 6. ถึง 11.
L3_QM ดูจาก ผลของ sheet เบอร์ 13. ถึง 16.
HW_Delta1.6 ดูจาก ผลของ sheet เบอร์ 18.
HW_Delta2.0 ดูจาก ผลของ sheet เบอร์ 19.
HW_Delta3.0 ดูจาก ผลของ sheet เบอร์ 20.


โดยถ้าเป็นกรณีที่พิจารณาจากหลายชีต ให้ เอา false เหนือ true ก็คือถ้าเจอว่าเป็น false ให้ return false ได้เลย
