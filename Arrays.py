class info:
    def __init__(self):
        self.Name = []
        self.Date = []
        self.Gen = []
        self.Address = []
        self.Phone = []
        self.Email = []
        self.Referred = []
        self.Prevdent = []
        self.Dentalcover = []
        self.Sensitivity = []
        self.Gum_Bleed = []
        self.Grind = []
        self.Cracking = []
        self.Emergencycon = []
        self.Emergencypho = []
        self.Relationship = []
        self.Physician = []
        self.Phyphone = []
        self.Medicondition = []
        self.Mediother =[]
        self.Sig = []
        self.Fax = []

    def up_date(self,date):
        if date =="":
            self.Date.append("NA")
        else:
            self.Date.append(date)

    def up_name(self,name):
        if name =="":
            self.Name.append("NA")
        else:
            self.Name.append(name)

    def up_gen(self,gen):
        if gen =="":
            self.Gen.append("NA")
        else:
            self.Gen.append(gen)

    def up_address(self,address):
        if address == "":
            self.Address.append("NA")
        else:
            self.Address.append(address)

    def up_phone(self,phone):
        if phone == "":
            self.Phone.append("NA")
        else:
            self.Phone.append(phone)
    def up_email(self,email):
        if email == "":
            self.Email.append("NA")
        else:
            self.Email.append(email)

    def up_referred(self,referred):
        if referred == "":
            self.Referred.append("NA")
        else:
            self.Referred.append(referred)

    def up_prevdent(self,prevdent):
        if prevdent == "":
            self.Prevdent.append("NA")
        else:
            self.Prevdent.append(prevdent)

    def up_dentalcover(self,dentalcover):
        if dentalcover == "":
            self.Dentalcover.append("NA")
        else:
            self.Dentalcover.append(dentalcover)

    def up_sensitivity(self,sensitivity):
        if sensitivity == "":
            self.Sensitivity.append("NA")
        else:
            self.Sensitivity.append(sensitivity)

    def up_gum_bleed(self,gum_bleed):
        if gum_bleed == "":
            self.Gum_Bleed.append("NA")
        else:
            self.Gum_Bleed.append(gum_bleed)

    def up_grind(self,grind):
        if grind == "":
            self.Grind.append("NA")
        else:
            self.Grind.append(grind)


    def up_cracking(self,cracking):
        if cracking == "":
            self.Cracking.append("NA")
        else:
            self.Cracking.append(cracking)

    def up_emergencycon(self,emergencycon):
        if emergencycon == "":
            self.Emergencycon.append("NA")
        else:
            self.Emergencycon.append(emergencycon)

    def up_emergencypho(self,emergencypho):
        if emergencypho == "":
            self.Emergencypho.append("NA")
        else:
            self.Emergencypho.append(emergencypho)

    def up_relationship(self,relationship):
        if relationship == "":
            self.Relationship.append("NA")
        else:
            self.Relationship.append(relationship)

    def up_physician(self,physician):
        if physician == "":
            self.Physician.append("NA")
        else:
            self.Physician.append(physician)

    def up_phyphone(self,phyphone):
        if phyphone == "":
            self.Phyphone.append("NA")
        else:
            self.Phyphone.append(phyphone)

    def up_medicondition(self,medicondition):
        if medicondition == "":
            self.Medicondition.append("NA")
        else:
            self.Medicondition.append(medicondition)
    def up_mediother(self,mediother):
        if mediother == "":
            self.Mediother.append("NA")
        else:
            self.Mediother.append(mediother)

    def up_sig(self,sig):
        if sig == "":
            self.Sig.append("NA")
        else:
            self.Sig.append(sig)

    def up_fax(self,fax):
        if fax == "":
            self.Fax.append("NA")
        else:
            self.Fax.append(fax)
