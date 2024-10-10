def tratar_string(string):
    # Encontrar a posição do terceiro underscore
    underscores = [pos for pos, char in enumerate(string) if char == '_']
    
    # Se houver pelo menos 3 underscores, retornar a parte após o terceiro
    if len(underscores) >= 3:
        return string[underscores[2] + 1:].strip()
    else:
        return string  # Retorna a string original se não houver 3 underscores

# Exemplos de uso
string1 = "TDE_SAP ECC_V1.4_ZEWM_Output ZEWM for shipment IDoc creation SHPMNT05_US - Copy"
string2 = "TDI_SAPECC_V1.0_Ghost 2.0_Interface 5_NF Issue Requisition - ID Creation"

print(tratar_string(string1))  # Saída: ZEWM_Output ZEWM for shipment IDoc creation SHPMNT05_US - Copy
print(tratar_string(string2))  # Saída: Ghost 2.0_Interface 5_NF Issue Requisition - ID Creation
