import { LightningElement } from 'lwc';
import { exportXLSX } from 'c/export_xlsx';

export default class Export_example_all extends LightningElement{

    handleExport(){
        let config = new Array();

        // function need to have @api notation
        config.push(this.template.querySelector('c-export_example_opportunity').getXLSConfig());
        
        // function need to have @api notation
        config.push(this.template.querySelector('c-export_example_products').getXLSConfig());

        // export one file with two worksheets
        exportXLSX(config, 'Opportunity and Products');
    }

}