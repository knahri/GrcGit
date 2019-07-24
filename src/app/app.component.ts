import { Component, OnInit, ɵConsole } from '@angular/core';
import docxtemplater from 'docxtemplater';
import * as JSZip from 'jszip';
import * as JSZipUtils from 'jszip-utils';
import { saveAs } from 'file-saver';
import { Input } from './models/input.model';
import { FormGroup, FormBuilder, FormArray } from '@angular/forms';


@Component({
  selector: 'app-root',

  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  titile = "";

  myForm: FormGroup;
  constructor(private evo: FormBuilder) { }

  ////////////////////////////////////////
  public livraison: string;
  public date: string;
  public cboxBiz: boolean = false;
  public cboxSrv: boolean = false;
  public cboxFx: boolean = false;
  public cboxGrcFx: boolean = false;
  public cboxMb: boolean = false;
  public cboxGrcMb: boolean = false;

  public IAplicatifMetier: string;
  public IAplicatifMetier2: string;

  public AAGrc: string;
  public AAGrc2: string;

  public installDA: string;
  public installDA2: string;

  public installGrcOpm: string;
  public installGrcOpm2: string;

  public BDAFx: String;
  public ScriptBDFx: String;

  public BDAMB: String;
  public ScriptBDMb: String;

  public DBA: String;
  public Arret_applicatif: String;
  public Script_BD: String;
  public Installation: String;

  public DDLFX: string;
  public DDLMB: string;

  public GRC_OPM_FX: any;
  public GRC_OPM_MB: any;


  public inputVal1: [];
  public inputVal2: [];
  public inputs: Input[] = [];



  /////////////// Check_Box ///////////////
  public cboxBizValue: any[] = [];
  onChangeCboxBiz() {
    if (this.cboxBiz) {
      this.cboxBizValue = [{
        "pieceBiz": "EAR GRC Business",
        "fichierBiz": "grcBusiness.ear",
        "expBiz": "EXP",
      }];
    } else {
      this.cboxBizValue = [];
    }
  }

  public cboxSrvValue: any[] = [];
  onChangeCboxSrv() {
    if (this.cboxSrv) {
      this.cboxSrvValue = [{
        "pieceSrv": "EAR GRC Services",
        "fichierSrv": "grcService.ear",
        "expSrv": "EXP",
      }];
    } else {
      this.cboxSrvValue = [];
    }
  }

  public cboxFxValue: any[] = [];
  onChangeCboxFx() {
    if (this.cboxFx) {
      this.cboxFxValue = [{
        "pieceFx": "LIVR_FX_" + this.livraison + "",
        "fichierFx": "GRC_BD_FX_" + this.livraison + ".sql",
        "expFx": "EXP",
      }];
    } else {
      this.cboxFxValue = [];
    }
  }

  public cboxMbValue: any[] = [];
  onChangeCboxMb() {
    if (this.cboxMb) {
      this.cboxMbValue = [{
        "pieceMb": "LIVR_MB_" + this.livraison + "",
        "fichierMb": "GRC_DB_MB_" + this.livraison + ".sql",
        "expMb": "EXP",
      }];
    } else {
      this.cboxMbValue = [];
    }
  }




  public cboxGrcFxValue: any[] = [];
  onChangeCboxGrcFx() {
    if (this.cboxGrcFx) {
      this.cboxGrcFxValue = [{
        "pieceGrcFx": "GRC_DBA_V" + this.livraison + "_FX",
        "fichierGrcFx": "DDL_GRC_FX_" + this.livraison + ".sql",
        "expGrcFx": "DBA",
      }];
    } else {
      this.cboxGrcFxValue = [];
    }
  }



  public cboxGrcMbValue: any[] = [];
  onChangeCboxGrcMb() {
    if (this.cboxGrcMb) {
      this.cboxGrcMbValue = [{
        "pieceGrcMb": "GRC_DBA_V" + this.livraison + "_MB",
        "fichierGrcMb": "DDL_GRC_MB_" + this.livraison + ".sql",
        "expGrcMb": "DBA",
      }];
    } else {
      this.cboxGrcMbValue = [];
    }
  }


  public OPM_MB: any[] = [];
  public OPM_FX: any[] = [];


  ////////////////////////////////////////////////////////
  loadFile(url) {
    JSZipUtils.getBinaryContent(url, (error, content) => {
      if (error) { throw error };
      var zip = new JSZip(content);

      var doc = new docxtemplater().loadZip(zip);
      ////////////////////////////////////////////////////////  
      if (this.cboxBiz === true) {
        this.IAplicatifMetier = "grcBusiness.ear"
      } else {
        this.IAplicatifMetier = "";
      }

      if (this.cboxSrv === true) {
        this.IAplicatifMetier2 = "grcService.ear"
      } else {
        this.IAplicatifMetier2 = "";
      }

      if (this.cboxBiz === true || this.cboxSrv === true) {
        this.AAGrc = "1.   GRC_OPM_Manuel_d_installation_GRC_V3.8.2.0_V01.0.docx [Exploitation]";
        this.AAGrc2 = "2.   GRC_OPM_Manual_d_indtallation_GRC_V3.8.2.0_V01.1.docx [Exploitation]";
        this.installDA = "1.   Déploiement applicatif [Exploitation]";
        this.installDA2 = "2.   Démarrage applicatif [Exploitation]";
        this.installGrcOpm = "  a.   GRC_OPM_Manuel_d_installation_GRC_V3.8.2.0_V01.0.docx ";
        this.installGrcOpm2 = "  a.   GRC_OPM_Manuel_d_installation_GRC_V3.8.2.0_V01.0.docx";
      } else {
        this.AAGrc = "";
        this.AAGrc2 = "";
        this.installDA = "";
        this.installDA2 = "";
        this.installGrcOpm = "";
        this.installGrcOpm2 = "";
      }

      ///////
      if (this.cboxFx === true) {
        this.BDAFx = "•	GRC_DBA_V3.8.2.0_FX.doc [DBA]"
      } else {
        this.BDAFx = "";
      }

      ///////
      if (this.cboxGrcFx === true) {
        this.ScriptBDFx = "1.	GRC_BD_FX_3.8.2.0.sql [Exploitation]"
      } else {
        this.ScriptBDFx = "";
      }

      //////
      if (this.cboxMb === true) {
        this.BDAMB = "•	GRC_DBA_V3.8.2.0_MB.doc [DBA]"
      } else {
        this.BDAMB = "";
      }

      //////
      if (this.cboxGrcMb === true) {
        this.ScriptBDMb = "2.	GRC_BD_MB_3.8.2.0.sql [Exploitation]"
      } else {
        this.ScriptBDMb = "";
      }


      if (this.cboxFx === true && this.cboxMb === true) {
        this.DBA = "DBA"
      } else {
        this.DBA = ""
      }

      if (this.cboxGrcFx === true && this.cboxGrcMb === true) {
        this.Script_BD = "Script BD"
      } else {
        this.Script_BD = ""
      }

      if (this.cboxBiz === true || this.cboxSrv === true) {
        this.Arret_applicatif = "Arrêt applicatif"
        this.Installation = "Installation"
      } else {
        this.Arret_applicatif = ""
        this.Installation = ""
      }

      if (this.cboxGrcMb === true) {
        this.DDLMB = "DDL_GRC_MB_" + this.livraison + ".sql"
      } else {
        this.DDLMB = ""
      }



      if (this.cboxGrcFx === true) {
        this.DDLFX = "DDL_GRC_FX_" + this.livraison + ".sql"
      } else {
        this.DDLFX = ""
      }

      ////////////////////////////////////////////////////////

  
      if (this.cboxMb === true) {
        this.OPM_MB = [{

        }];
      } else {
        this.OPM_MB = [];
      }

      if (this.cboxFx === true) {
        this.OPM_FX = [{

        }];
      } else {
        this.OPM_FX = [];
      }
      
      ////////////////////////////////////////////////////////

      doc.setData({

        "version": this.livraison,
        "date": this.date,

        "varBiz": this.cboxBizValue,
        "varSrv": this.cboxSrvValue,
        "varFx": this.cboxFxValue,
        "varGrcFx": this.cboxGrcFxValue,

        "varMb": this.cboxMbValue,
        "varGrcMb": this.cboxGrcMbValue,

        "IApplicatifMetier": this.IAplicatifMetier,
        "IApplicatifMetier2": this.IAplicatifMetier2,

        "AAGrc": this.AAGrc,
        "AAGrs2": this.AAGrc2,

        "installDA": this.installDA,
        "installDA2": this.installDA2,
        "installGrcOpm": this.installGrcOpm,
        "installGrcOpm2": this.installGrcOpm2,
        "instalgrcopm": this.IAplicatifMetier2,

        "DBAFX": this.BDAFx,
        "ScriptBDFx": this.ScriptBDFx,

        "DBAMB": this.BDAMB,
        "ScriptBDMb": this.ScriptBDMb,

        "DBA": this.DBA,
        "Script BD": this.Script_BD,
        "Arrêt applicatif": this.Arret_applicatif,
        "Installation": this.Installation,
        "fichierGrcMb": this.DDLMB,
        "fichierGrcFx": this.DDLFX,

        "OPM_MB": this.OPM_MB,
        "OPM_FX": this.OPM_FX,
   


      });
      ///////////////////////////////////////////////////

      doc.render()
      var out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      })
      //saveAs(out, 'GRC_V' + this.livraison + '_Dossier-Livraison_V1.0.docx');
      //saveAs(out, 'GRC_FDV_Livraison_VX.X.X.X.docx');
      //saveAs(out, 'GRC_DBA_V3.8.2.0.0_FX.docx');
      //saveAs(out, 'GRC_DBA_V3.8.2.0.0_MB.docx');
      saveAs(out, 'GRC_OPM_MB.docx');
      saveAs(out, 'GRC_OPM_FX.docx');
    })
  }
  /*
  generate1() {
    this.loadFile("assets/GRC_VX.X.X.X_Dossier-Livraison_V1.0.docx");
  }
  generate2() {
    this.loadFile('assets/GRC_FDV_Livraison_VX.X.X.X.docx');
  }
  generate3() {
    this.loadFile('assets/GRC_DBA_V3.8.2.0.0_FX.docx');
  }
  generate4() {
    this.loadFile('assets/GRC_DBA_V3.8.2.0.0_MB.docx');
  }
  */
  generate5() {
    this.loadFile('assets/GRC_OPM_MB.docx');
  }
  generate6() {
    this.loadFile('assets/GRC_OPM_FX.docx');
  }


  call() {
    //this.generate1();
    //this.generate2();
    //this.generate3();
    //this.generate4();
    this.generate5();
    this.generate6();
    
  }
  ///////////////////////////////////////////////////////////////

  ngOnInit() {
    this.myForm = this.evo.group({
      id2: '',
      evol2: this.evo.array([]),
    })
  }

  get diForms() {
    return this.myForm.get('evol2') as FormArray

  }

  ajouter() {
    const constid = this.evo.group({
      id: [],
      evol: [],
    })
    this.diForms.push(constid);
  }

  supprimer(i) {
    this.diForms.removeAt(i)
  }
}








/*
  enregistrer() {
    this.i = 0;
    this.donne = "";
    this.donne2 = "";
    do {
      this.donne = this.donne.concat("\n", this.inputs[this.i].inputVal1);
      this.donne2 = this.donne2.concat("\n", this.inputs[this.i].inputVal2);
      this.i++;
    } while (this.i < this.inputs.length);
  }


  ajouterT() {
    this.inputs.push(new Input());

  }
  getInputsVal() {
    console.log('get inputs val ', this.inputs)
    //this.enregistrer()
    console.log('donne ', this.donne, 'donne2 ', this.donne2)
  }
*/
  /*
  ajouterT() {
    this.inputs.push(new Input());
  }


  getInputsVal() {
    console.log('get inputs val ', this.inputs)
  }

             "ng": [{
             //   "ngid": "" + this.ngidV,
             //   "ngevol": "" + this.ngevolV + "",

             "ngid": "" + this.inputs[0].inputVal1 + "",
             "ngevol": "" + this.inputs[0].inputVal2 + "",
           }]
*/