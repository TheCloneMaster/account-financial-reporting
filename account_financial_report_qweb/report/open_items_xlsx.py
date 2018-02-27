# -*- coding: utf-8 -*-
# Author: Julien Coux
# Copyright 2016 Camptocamp SA
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

from . import abstract_report_xlsx
from odoo.report import report_sxw
from odoo import _
from datetime import date

class OpenItemsXslx(abstract_report_xlsx.AbstractReportXslx):

    def __init__(self, name, table, rml=False, parser=False, header=True,
                 store=False):
        super(OpenItemsXslx, self).__init__(
            name, table, rml, parser, header, store)

    def _get_report_name(self):
        return _('Open Items')

    def _get_report_columns(self, report):
        return {
            0: {'header': _('Fecha'), 'field': 'date', 'width': 11},
            1: {'header': _('Factura'), 'field': 'entry', 'width': 18},
            2: {'header': _('Vencimiento'), 'field': 'date_due', 'width': 11},
            
            3: {'header': _('Plazo'), 'field': 'term', 'width': 8},
            4: {'header': _('Días Vencidos'), 'field': 'due_days', 'type': 'amount','width': 9},
            5: {'header': _('Total Factura ¢'),
                'field': 'amount_total_due',
                'type': 'amount',
                'width': 14},
            6: {'header': _('Saldo ¢'),
                'field': 'amount_residual',
                'field_final_balance': 'cumulative_crc',
                'type': 'amount',
                'width': 14},
            7: {'header': _('Total Factura $'),
                 'field': 'amount_total_due_currency',
                 'type': 'amount',
                 'width': 14},
            8: {'header': _('Saldo $'),
                'field': 'amount_residual_currency',
                'field_final_balance': 'cumulative_usdx',
                'type': 'amount',
                'width': 14},
            9: {'header': _('Acum. Colones'),
                 'field': 'cumulative',
                 'type': 'amount',
                 'width': 14},
            10: {'header': _('Acum. Dólares'),
                 'field': 'cumulative_usd',
                 'type': 'amount',
                 'width': 14},
        }
        ################  ORIGINAL VERSION
        return {
            0: {'header': _('Date'), 'field': 'date', 'width': 11},
            1: {'header': _('Entry'), 'field': 'entry', 'width': 18},
            2: {'header': _('Journal'), 'field': 'journal', 'width': 8},
            3: {'header': _('Account'), 'field': 'account', 'width': 9},
            4: {'header': _('Partner'), 'field': 'partner', 'width': 25},
            5: {'header': _('Ref - Label'), 'field': 'label', 'width': 40},
            6: {'header': _('Due date'), 'field': 'date_due', 'width': 11},
            7: {'header': _('Original'),
                'field': 'amount_total_due',
                'type': 'amount',
                'width': 14},
            8: {'header': _('Residual'),
                'field': 'amount_residual',
                'field_final_balance': 'final_amount_residual',
                'type': 'amount',
                'width': 14},
            9: {'header': _('Cur.'), 'field': 'currency_name', 'width': 7},
            10: {'header': _('Cur. Original'),
                 'field': 'amount_total_due_currency',
                 'type': 'amount',
                 'width': 14},
            11: {'header': _('Cur. Residual'),
                 'field': 'amount_residual_currency',
                 'type': 'amount',
                 'width': 14},
        }


    def _get_report_filters(self, report):
        return [
            [_('Date at filter'), report.date_at],
            [_('Target moves filter'),
                _('All posted entries') if report.only_posted_moves
                else _('All entries')],
            [_('Account balance at 0 filter'),
                _('Hide') if report.hide_account_balance_at_0 else _('Show')],
        ]

    def _get_col_count_filter_name(self):
        return 2

    def _get_col_count_filter_value(self):
        return 2

    def _get_col_count_final_balance_name(self):
        return 5

    def _get_col_pos_final_balance_label(self):
        return 5

    def _generate_report_content(self, workbook, report):
        # For each account
        for account in report.account_ids:
            # Write account title
            self.write_array_title(account.code + ' - ' + account.name)
            acc_cumulative = acc_cumulative_usd = 0
            # For each partner
            for partner in account.partner_ids:
                # Write partner title
                self.write_array_title(partner.name)

                # Display array header for move lines
                self.write_array_header()
                
                cumulative = cumulative_usd = 0

                # Display account move lines
                for line in partner.move_line_ids:
                    if (line.amount_residual_currency and abs(line.amount_residual_currency)<1.0) or abs(line.amount_residual)<1.0:
                        continue
                    elif line.amount_residual_currency:
                        cumulative_usd += line.amount_residual_currency
                        setattr(line, 'amount_residual', 0.0)
                    else:
                        cumulative += line.amount_residual
                    setattr(line, 'cumulative', cumulative)
                    setattr(line, 'cumulative_usd', cumulative_usd)
                    #setattr(line, 'term', partner.partner_id.property_payment_term_id.name)
                    if line.move_line_id.invoice_id and line.move_line_id.invoice_id.payment_term_id:
                        setattr(line, 'term', line.move_line_id.invoice_id.payment_term_id.name)
                    else:
                        setattr(line, 'term', '')
                    setattr(line, 'due_days', (date.today()-
                    date(int(line.date_due[0:4]), int(line.date_due[5:7]), int(line.date_due[8:10]))).days)
                    self.write_line(line)

                class partner_balance():
                    cumulative_crc = cumulative
                    cumulative_usdx = cumulative_usd
                    name = partner.name
                    
                # Display ending balance line for partner
                self.write_ending_balance(partner_balance, 'partner')

                # Line break
                self.row_pos += 1
                acc_cumulative += cumulative
                acc_cumulative_usd += cumulative_usd

            class account_balance():
                cumulative_crc = acc_cumulative
                cumulative_usdx = acc_cumulative_usd
                name = account.name
                code = account.code
            # Display ending balance line for account
            self.write_ending_balance(account_balance, 'account')

            # 2 lines break
            self.row_pos += 2

    def write_ending_balance(self, my_object, type_object):
        """Specific function to write ending balance for Open Items"""
        if type_object == 'partner':
            name = my_object.name
            label = _('Partner ending balance')
        elif type_object == 'account':
            name = my_object.code + ' - ' + my_object.name
            label = _('Ending balance')
        super(OpenItemsXslx, self).write_ending_balance(my_object, name, label)


OpenItemsXslx(
    'report.account_financial_report_qweb.report_open_items_xlsx',
    'report_open_items_qweb',
    parser=report_sxw.rml_parse
)
