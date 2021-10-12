import boto3
import re
import sys

from xl_helper import ExcelSheet


def camel_to_space(name):
    name = re.sub('(.)([A-Z][a-z]+)', r'\1 \2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1 \2', name)


def get_savings_plans_recommendations(profile, term="ONE_YEAR", payment_options="NO_UPFRONT", account_scope="PAYER", look_back_period='THIRTY_DAYS'):
    # 'AccountScope': 'PAYER'|'LINKED',
    # 'SavingsPlansType': 'COMPUTE_SP'|'EC2_INSTANCE_SP'|'SAGEMAKER_SP',
    # TermInYears='ONE_YEAR'|'THREE_YEARS'
    # PaymentOptions: 'NO_UPFRONT'|'PARTIAL_UPFRONT'|'ALL_UPFRONT'|'LIGHT_UTILIZATION'|'MEDIUM_UTILIZATION'|'HEAVY_UTILIZATION'
    # AccountScope: 'PAYER'|'LINKED'
    # LookbackPeriodInDays='SEVEN_DAYS'|'THIRTY_DAYS'|'SIXTY_DAYS',
    #
    profile_session = boto3.session.Session(profile_name=profile)
    ce_client = profile_session.client('ce')
    ce_savings_plans = ce_client.get_savings_plans_purchase_recommendation(
        SavingsPlansType='COMPUTE_SP',
        TermInYears=term,
        PaymentOption=payment_options,
        AccountScope=account_scope,
        LookbackPeriodInDays=look_back_period,
    )
    # print(ce_savings_plans)
    if 'SavingsPlansPurchaseRecommendationDetails' in ce_savings_plans['SavingsPlansPurchaseRecommendation']:
        alias = profile_session.client(
            'iam').list_account_aliases()['AccountAliases'][0]
        savings_plans = ce_savings_plans['SavingsPlansPurchaseRecommendation']['SavingsPlansPurchaseRecommendationDetails'][0]
        sp = {**{'AccountAliases': alias, 'Term': term}, **savings_plans}
        return sp
    else:
        return None


def is_float(value):
    try:
        float(value)
        return True
    except ValueError:
        return False


def convert_to_number(number_string):
    """
    convert a string to either int or float
    """
    if number_string.isdigit():
        return int(number_string)
    elif is_float(number_string):
        return float(number_string)
    else:
        return number_string


def write_sp_to_excel(xl):
    """
    Write the output from Savings Plans to Excel xlsx 
    """
    profiles = boto3.session.Session().available_profiles
    worksheet_name = 'Savings Plans'
    xl.add_worksheet(worksheet_name)
    terms = ['ONE_YEAR', 'THREE_YEARS']
    ignore_profiles = ['default', 'Billing']
    for profile in profiles:
        print(f'Getting Savings plans for: {profile}')
        if profile in ignore_profiles:
            continue

        for term in terms:
            sp_rec = get_savings_plans_recommendations(
                profile, term=term, account_scope='LINKED')
            if not sp_rec:
                continue

            headers = []
            values = []
            formats = [xl.PLAIN, xl.PLAIN, xl.NUMBER, xl.CURRENCY, xl.CURRENCY, xl.PLAIN, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY,
                       xl.CURRENCY, xl.DECIMAL, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY]
            for k, v in sp_rec.items():
                if not isinstance(v, dict):
                    # add this value to Excel worksheet with space instead of CamelCase
                    headers.append(camel_to_space(k))
                    # see of value is int or float and then add to array
                    values.append(convert_to_number(v))

            xl.add_header_row(worksheet_name, headers)
            xl.add_row(worksheet_name, values, formats)
            xl.add_conditional_format_column(worksheet_name, 10)
            xl.add_autofilter(worksheet_name, len(values))

def get_reservation_recommendations(profile, service, term='ONE_YEAR', payment_options="NO_UPFRONT", look_back_period='THIRTY_DAYS'):
    """
    Fetch the Cost Explorer reservation recommendations
    """
    profile_session = boto3.session.Session(profile_name=profile)
    ce_client = profile_session.client('ce')
    response = {}
    try:
        response = ce_client.get_reservation_purchase_recommendation(
                LookbackPeriodInDays=look_back_period,
                TermInYears=term,
                PaymentOption=payment_options,
                Service=service
            )
    except Exception as e:
        print(f'Error gathering recommendations for {service}: {str(e)}') 
        sys.exit(1)
    recommendations = response['Recommendations']
    if recommendations:
        ri_rec = recommendations[0]
        alias = profile_session.client(
                'iam').list_account_aliases()['AccountAliases'][0]
        ri_rec = {**{'AccountAliases': alias}, **ri_rec}

        return ri_rec
    else: 
        return None

def write_ri_to_excel(xl):
    """
    Write the out put from AWS Cost Explorer Reserved Instances to xlsx 
    """
    profiles = boto3.session.Session().available_profiles
    worksheet_name = 'Reserved Instances'
    xl.add_worksheet(worksheet_name)
    terms = ['ONE_YEAR', 'THREE_YEARS']
    services = [
        'Amazon Elastic Compute Cloud - Compute', 'Amazon Relational Database Service', 
    ]
    # services not checked: 
    # 'Amazon Redshift', 'Amazon ElastiCache', 'Amazon Elasticsearch Service'
    ignore_profiles = ['default', 'Billing']
    for profile in profiles:
        if profile in ignore_profiles:
            continue
        for service in services:
            for term in terms:
                
                print(f'Working on {service} and term {term}...')
                ri_recommendations = get_reservation_recommendations(
                    profile, service, term=term)
                if not ri_recommendations:
                    continue
                prefix_headers = []
                prefix_values = []
                formats = [xl.PLAIN, xl.PLAIN, xl.NUMBER, xl.NUMBER, xl.PLAIN, # Prefix headers
                        xl.NUMBER, xl.PLAIN, xl.PLAIN, xl.NUMBER, xl.NUMBER, xl.NUMBER, xl.NUMBER, xl.NUMBER, xl.NUMBER,
                        xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.NUMBER, xl.NUMBER, xl.PLAIN, xl.CURRENCY, xl.DECIMAL, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY]
                for k, v in ri_recommendations.items():
                    if isinstance(v, list) and 'RecommendationDetails' in k:
                        ri_rec_details = ri_recommendations['RecommendationDetails']
                        for rec in ri_rec_details:
                            headers = []
                            values = []
                            for key, value in rec.items():
                                if 'InstanceDetails' in key:
                                    instance_details = list(value.keys())[0]
                                    headers.append('InstanceType')
                                    values.append(value[instance_details]['InstanceType'])
                                    headers.append('Region')
                                    values.append(value[instance_details]['Region'])
                                else:
                                    # add this value to Excel worksheet with space instead of CamelCase
                                    headers.append(camel_to_space(key))
                                    # see of value is int or float and then add to array
                                    values.append(convert_to_number(value))
                            xl.add_header_row(worksheet_name, prefix_headers + ['Service'] + headers)
                            xl.add_row(worksheet_name, prefix_values + [service] + values, formats)
                            xl.add_conditional_format_column(worksheet_name, 21)
                            xl.add_autofilter(worksheet_name, len(values))

                    elif not isinstance(v, dict):
                        # add this value to Excel worksheet with space instead of CamelCase
                        prefix_headers.append(camel_to_space(k))
                        # see of value is int or float and then add to array
                        prefix_values.append(convert_to_number(v))

if __name__ == "__main__":
    file_name_prefix = 'CostSavingRecommendations'
    xl = ExcelSheet(file_name_prefix)

    write_sp_to_excel(xl)
    write_ri_to_excel(xl)
    xl.close()
