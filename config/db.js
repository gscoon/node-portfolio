var mysql = require('mysql');

var dbClass = function(){
    console.log('MYSQL obj created');
    this.conn = mysql.createConnection({
        host: site.config.db.server,
        port: site.config.db.port,
        user: site.config.db.user,
        password: site.config.db.password,
        database: site.config.db.name
    });

    this.getAnalysisTemplateData = function(mfiID, callBack){
        var q = [];
        var params = [mfiID];

        // monthy financials
        q.push("SELECT DS.sheet_name, FL.field_label_english, FS.is_audited, FS.field_set_date, F.field_value FROM reported_financials F LEFT JOIN reporting_field_set FS ON FS.field_set_id = F.field_set_id LEFT JOIN mfi M ON (M.mfi_id = FS.mfi_id) LEFT JOIN reporting_field_label FL ON F.field_label_id = FL.field_label_id LEFT JOIN reporting_data_sheet DS ON DS.sheet_id = FS.sheet_id WHERE M.mfi_id = ?");

        //histroical FX rates
        q.push("SELECT concat(MONTH(FX.fx_date),'-',YEAR(FX.fx_date)), FX.conversion_currency_code, FX.fx_value, FX.fx_date FROM (SELECT FX.conversion_currency_code, MAX(FX.fx_date) as mfd FROM fx_historical FX WHERE FX.base_currency_code = 'USD' GROUP BY fx.conversion_currency_code, YEAR(FX.fx_date), MONTH(FX.fx_date) ) MF LEFT JOIN fx_historical FX ON (FX.conversion_currency_code = MF.conversion_currency_code AND FX.fx_date = MF.mfd AND FX.base_currency_code = 'USD') LEFT JOIN mfi M ON (M.reporting_currency = FX.conversion_currency_code) WHERE M.mfi_id = ? ORDER BY FX.fx_date DESC");

        //sum of transactions for each fund
        q.push("SELECT f.fund_name, st.sumt as 'Sum of Deals (OC)', st.sumt_fx as 'Sum of Deals ($ at Transaction Date)',  (st.sumt/fx.fx_value) as 'Sum of Deals ($ at Latest FX Date)', st.currency_code as 'Deal Currency', st.deal_date, st.d_mat as 'Deal Maturity', if(st.index_name IS NULL, st.crate, concat(round(st.crate, 2), '% ', st.plus_or_minus_rate, ' ',st.index_name)) as 'Coupon Rate', st.deal_type_id as 'Deal Type' FROM (SELECT d.mfi_id, d.deal_date, d.deal_type_id, t.fund_id, t.currency_code, sum(t.transaction_amount) as sumt, sum(t.transaction_amount * 1/fx.fx_value) as sumt_fx, max(dd.maturity_date) as d_mat, max(dd.coupon_rate) as crate, IRI.index_name, dd.plus_or_minus_rate FROM mfi_deal d LEFT JOIN mfi_deal_transaction t ON t.deal_id = d.deal_id LEFT JOIN fx_historical fx ON (fx.fx_date = T.transaction_date AND fx.conversion_currency_code = t.currency_code AND fx.base_currency_code = 'USD') LEFT JOIN disbursement_detail_debt dd ON dd.transaction_id = t.transaction_id LEFT JOIN interest_rate_index IRI ON IRI.index_id = dd.interest_rate_index_id GROUP BY d.mfi_id, d.deal_date, t.fund_id, t.currency_code HAVING sum(t.transaction_amount) IS NOT NULL) st LEFT JOIN mfi m ON m.mfi_id = st.mfi_id LEFT JOIN fund f ON f.fund_id = st.fund_id LEFT JOIN (SELECT max(fx_date) as last_date, conversion_currency_code FROM fx_historical WHERE base_currency_code = 'USD' GROUP BY conversion_currency_code)  max_fx ON max_fx.conversion_currency_code = st.currency_code LEFT JOIN fx_historical fx ON fx.base_currency_code = 'USD' AND fx.conversion_currency_code = max_fx.conversion_currency_code AND fx.fx_date = max_fx.last_date WHERE m.mfi_id = ? AND st.sumt > 0 ORDER BY st.d_mat ASC, f.fund_name");

        //amortization history
        q.push("SELECT CONCAT(st.d_mat,'-',f.fund_name) as 'lookup_key', f.fund_name as 'Fund', st.sumt as 'Sum of Deals (OC)', (st.sumt/fx.fx_value) as 'Sum of Deals ($ at Latest FX Date)', st.currency_code as 'Deal Currency', st.d_mat as 'Tranche Maturity' FROM (SELECT d.mfi_id, d.deal_date, d.deal_type_id, t.fund_id, t.currency_code, sum(t.transaction_amount) as sumt, sum(t.transaction_amount * 1/fx.fx_value) as sumt_fx, max(dd.maturity_date) as d_mat, max(dd.coupon_rate) as crate, IRI.index_name, dd.plus_or_minus_rate FROM mfi_deal d LEFT JOIN mfi_deal_transaction t ON t.deal_id = d.deal_id LEFT JOIN fx_historical fx ON (fx.fx_date = T.transaction_date AND fx.conversion_currency_code = t.currency_code AND fx.base_currency_code = 'USD') RIGHT JOIN disbursement_detail_debt dd ON dd.transaction_id = t.transaction_id LEFT JOIN interest_rate_index IRI ON IRI.index_id = dd.interest_rate_index_id GROUP BY d.mfi_id, t.fund_id, t.currency_code, dd.maturity_date HAVING sum(t.transaction_amount) IS NOT NULL) st LEFT JOIN mfi m ON m.mfi_id = st.mfi_id LEFT JOIN fund f ON f.fund_id = st.fund_id LEFT JOIN (SELECT max(fx_date) as last_date, conversion_currency_code FROM fx_historical WHERE base_currency_code = 'USD' GROUP BY conversion_currency_code)  max_fx ON max_fx.conversion_currency_code = st.currency_code LEFT JOIN fx_historical fx ON fx.base_currency_code = 'USD' AND fx.conversion_currency_code = max_fx.conversion_currency_code AND fx.fx_date = max_fx.last_date WHERE m.mfi_id = ? AND st.sumt > 0 ORDER BY st.d_mat ASC, f.fund_name, st.currency_code");

        //covenant information
        q.push("SELECT ct.covenant_type, c.covenant_value, c.start_date, c.end_date FROM mfi_covenant c LEFT JOIN mfi_covenant_type ct ON ct.covenant_type_id = c.covenant_type_id WHERE c.mfi_id=?");

        //mfi commentary
        q.push("SELECT c.comment_date,  c.comments FROM mfi_commentary c WHERE c.mfi_id = ? ORDER BY c.comment_date DESC");

        // rating
        q.push("SELECT r.rating_date, rl.rating_label_short, rl.rating_label_long FROM mfi_risk_rating r LEFT JOIN mfi_risk_rating_label rl ON r.rating_label_id = rl.rating_label_id WHERE r.mfi_id = ? ORDER BY r.rating_date DESC");

        // country mfis deals
        q.push("SELECT F.fund_name, M.mfi_name, M.mfi_id, C.country_name, st.sumt, st.sumt_fx, st.currency_code, st.deal_type_id, (st.sumt/fx.fx_value) as 'Amount as of Last FX'  FROM (SELECT d.mfi_id, d.deal_date, d.deal_type_id, t.fund_id, t.currency_code, sum(t.transaction_amount) as sumt, sum(t.transaction_amount * 1/fx.fx_value) as sumt_fx, max(dd.maturity_date) as d_mat, max(dd.coupon_rate) as crate FROM mfi_deal d LEFT JOIN mfi_deal_transaction t ON t.deal_id = d.deal_id LEFT JOIN fx_historical fx ON (fx.fx_date = T.transaction_date AND fx.conversion_currency_code = t.currency_code AND fx.base_currency_code = 'USD') LEFT JOIN disbursement_detail_debt dd ON dd.transaction_id = t.transaction_id GROUP BY d.mfi_id, d.deal_date, t.fund_id, t.currency_code HAVING sum(t.transaction_amount) IS NOT NULL) st LEFT JOIN mfi M ON st.mfi_id = M.mfi_id LEFT JOIN fund f ON f.fund_id = st.fund_id LEFT JOIN country C ON C.country_id = M.country_id LEFT JOIN (SELECT max(fx_date) as last_date, conversion_currency_code FROM fx_historical WHERE base_currency_code = 'USD' GROUP BY conversion_currency_code) max_fx ON max_fx.conversion_currency_code = st.currency_code LEFT JOIN fx_historical fx ON fx.base_currency_code = 'USD' AND fx.conversion_currency_code = max_fx.conversion_currency_code AND fx.fx_date = max_fx.last_date WHERE M.country_id IN (SELECT country_id FROM mfi WHERE mfi_id = ?) AND st.sumt > 0");

        // social data
        q.push("SELECT m.mfi_name, m.mfi_id, sf.field_name, sv.field_id, sv.field_value, ss.questionnaire_year FROM reporting_social_value sv JOIN reporting_social_field sf ON sf.field_id = sv.field_id JOIN reporting_social_submission ss ON ss.submission_id = sv.submission_id JOIN mfi m ON m.mfi_id = ss.mfi_id LEFT JOIN (SELECT mfi_id, MAX(questionnaire_year) as max_yr FROM reporting_social_submission GROUP BY mfi_id) max_ss ON max_ss.mfi_id = m.mfi_id WHERE sf.field_id = 353 AND ss.questionnaire_year = max_ss.max_yr AND m.mfi_id = ?");

        // other ratings
        q.push("SELECT rating.rating_type, rating.rating_date, rating.rating_value FROM mfi_other_rating rating JOIN mfi m ON m.mfi_id = rating.mfi_id LEFT JOIN (SELECT mfi_id, rating_type, MAX(rating_date) as max_date FROM mfi_other_rating GROUP BY mfi_id, rating_type) max_rating ON max_rating.mfi_id = rating.mfi_id AND max_rating.rating_type = rating.rating_type WHERE rating.rating_date = max_rating.max_date AND rating.mfi_id = ?");

        site.async.map(q, function(item, mapCallBack){
            runQuery(item, params, mapCallBack);
        }, callBack);

    }

    function runQuery(q, params, callback){
        site.db.conn.query(q, params, callback);
    }

}

module.exports = new dbClass();
