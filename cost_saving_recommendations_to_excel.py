import boto3
import re
import sys
import argparse

from xl_helper import ExcelSheet


def camel_to_space(name):
    name = re.sub('(.)([A-Z][a-z]+)', r'\1 \2', name)
    return re.sub('([a-z0-9])([A-Z])', r'\1 \2', name)


def get_savings_plans_recommendations(profile, term="ONE_YEAR", payment_options="NO_UPFRONT", account_scope="PAYER", look_back_period='THIRTY_DAYS', sp_type = 'COMPUTE_SP'):
    # 'AccountScope': 'PAYER'|'LINKED',
    # 'SavingsPlansType': 'COMPUTE_SP'|'EC2_INSTANCE_SP'|'SAGEMAKER_SP',
    # TermInYears='ONE_YEAR'|'THREE_YEARS'
    # PaymentOptions: 'NO_UPFRONT'|'PARTIAL_UPFRONT'|'ALL_UPFRONT'|'LIGHT_UTILIZATION'|'MEDIUM_UTILIZATION'|'HEAVY_UTILIZATION'
    # AccountScope: 'PAYER'|'LINKED'
    # LookbackPeriodInDays='SEVEN_DAYS'|'THIRTY_DAYS'|'SIXTY_DAYS',
    #
    profile_session = boto3.session.Session(profile_name=profile)
    ce_client = profile_session.client('ce')

    print('Running command:')
    print(f'aws ce get-savings-plans-purchase-recommendation --savings-plans-type "{sp_type}" --term-in-years "{term}" --payment-option "{payment_options}" --lookback-period-in-days "{look_back_period}" --profile "{profile}"')

    ce_savings_plans = ce_client.get_savings_plans_purchase_recommendation(
        SavingsPlansType=sp_type,
        TermInYears=term,
        PaymentOption=payment_options,
        AccountScope=account_scope,
        LookbackPeriodInDays=look_back_period,
    )
    if 'SavingsPlansPurchaseRecommendationDetails' in ce_savings_plans['SavingsPlansPurchaseRecommendation']:
        alias = profile_session.client(
            'iam').list_account_aliases()['AccountAliases'][0]
        savings_plans = ce_savings_plans['SavingsPlansPurchaseRecommendation']['SavingsPlansPurchaseRecommendationDetails'][0]
        sp = {**{'AccountAliases': alias, 'Term': term, 'PaymentOption': payment_options}, **savings_plans}
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


def write_sp_to_excel(xl, terms = ['ONE_YEAR', 'THREE_YEARS'], payment_options = ['NO_UPFRONT']):
    """
    Write the output from Savings Plans to Excel xlsx 
    Loop through profiles, payment options and terms in years
    """
    profiles = boto3.session.Session().available_profiles
    worksheet_name = 'Savings Plans'
    xl.add_worksheet(worksheet_name)
    ignore_profiles = ['default', 'Billing']
    for profile in profiles:
        print(f'Getting Savings plans for: {profile}')
        if profile in ignore_profiles:
            continue
        for po in payment_options:
            for term in terms:
                sp_rec = get_savings_plans_recommendations(
                    profile, term=term, payment_options=po , account_scope='LINKED')
                if not sp_rec:
                    continue

                headers = []
                values = []
                formats = [xl.PLAIN, xl.PLAIN, xl.PLAIN, xl.NUMBER, xl.CURRENCY, xl.CURRENCY, xl.PLAIN, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY,
                        xl.CURRENCY, xl.DECIMAL, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY]
                for k, v in sp_rec.items():
                    if not isinstance(v, dict):
                        # add this value to Excel worksheet with space instead of CamelCase
                        headers.append(camel_to_space(k))
                        # see of value is int or float and then add to array
                        values.append(convert_to_number(v))

                xl.add_header_row(worksheet_name, headers)
                xl.add_row(worksheet_name, values, formats)
                index = [idx for idx, s in enumerate(headers) if 'Persent' in s][0]
                xl.add_conditional_format_column(worksheet_name, index)
                xl.add_autofilter(worksheet_name, len(values))

def get_reservation_recommendations(profile, service, term='ONE_YEAR', payment_options="NO_UPFRONT", look_back_period='THIRTY_DAYS'):
    """
    Fetch the Cost Explorer reservation recommendations
    """
    profile_session = boto3.session.Session(profile_name=profile)
    ce_client = profile_session.client('ce')
    response = {}
    try:
        print('Running command: ')
        print(f'aws ce get-reservation-purchase-recommendation --service "{service}" --term-in-years "{term}" --payment-option "{payment_options}" --lookback-period-in-days "{look_back_period}" --profile "{profile}"')
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
        ri_recommendations_details = recommendations[0]['RecommendationDetails']
        alias = profile_session.client(
                'iam').list_account_aliases()['AccountAliases'][0]
        prefix_details = {'AccountAliases': alias, 'TermInYears': term, 'PaymentOption': payment_options}

        return (prefix_details, ri_recommendations_details)
    else: 
        return None

def write_ri_to_excel(xl, terms = ['ONE_YEAR', 'THREE_YEARS'], payment_options=['NO_UPFRONT']):
    """
    Write the out put from AWS Cost Explorer Reserved Instances to xlsx 
    For all availabe profiles in aws config
    """
    profiles = boto3.session.Session().available_profiles
    worksheet_name_prefix = 'RI'
    services = [
        'Amazon Elastic Compute Cloud - Compute', 'Amazon Relational Database Service', 
        # services not used here: 
        # 'Amazon Redshift', 'Amazon ElastiCache', 'Amazon Elasticsearch Service'
    ]
    # worksheet names are restricted to less than 31 char, we need to create a short version
    services_short = dict(zip(services, ['EC2', 'RDS']))
    # add one worksheet per service
    for service in services:
        xl.add_worksheet(f'{worksheet_name_prefix} - {services_short[service]}')

    ignore_profiles = ['default', 'Billing']
    for profile in profiles:
        if profile in ignore_profiles:
            continue
        for service in services:
            # write data to the active worksheet
            worksheet_name = f'{worksheet_name_prefix} - {services_short[service]}'
            for po in payment_options:
                for term in terms:
                    ri_recommendations = get_reservation_recommendations(
                        profile, service, term=term, payment_options=po)
                    if not ri_recommendations:
                        continue
                    prefix_headers = [camel_to_space(k) for k in ri_recommendations[0].keys()]
                    prefix_values = list(ri_recommendations[0].values())
                    formats = [xl.PLAIN, xl.PLAIN, xl.NUMBER, xl.NUMBER, xl.PLAIN, # Prefix headers
                            xl.NUMBER, xl.PLAIN, xl.PLAIN, xl.NUMBER, xl.NUMBER, xl.NUMBER, xl.NUMBER, xl.NUMBER, xl.CURRENCY, xl.CURRENCY,
                            xl.CURRENCY, xl.NUMBER, xl.NUMBER, xl.PLAIN, xl.CURRENCY, xl.DECIMAL, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY, xl.CURRENCY]
                    for ri in ri_recommendations[1]:
                        instance_details = ri.pop('InstanceDetails')
                        instance_details = list(instance_details.values())[0]
                        headers = [camel_to_space(k) for k in ri.keys()]
                        values = list(ri.values())
                        for id in ['InstanceType', 'Region']:
                            headers.append(id)
                            values.append(instance_details[id])
                        xl.add_header_row(worksheet_name, prefix_headers + headers)
                        xl.add_row(worksheet_name, prefix_values + values, formats)
                        index = [idx for idx, s in enumerate(headers) if 'Persent' in s][0]
                        xl.add_conditional_format_column(worksheet_name, index)
                        xl.add_autofilter(worksheet_name, len(values))
                        

if __name__ == "__main__":
    # parse cli arguments
    parser = argparse.ArgumentParser()
    parser.add_argument('-p', '--payment-options',
                        help="Payment Options: 'NO_UPFRONT'|'PARTIAL_UPFRONT'|'ALL_UPFRONT' or several with comma [,] separation",
                        default="NO_UPFRONT,PARTIAL_UPFRONT",
                        required=False,
                        dest='payment_options')    
    parser.add_argument('-t', '--term-in-years',
                        help="Term In Years 'ONE_YEAR'|'THREE_YEARS' or both separated with comma [,]",
                        default="ONE_YEAR,THREE_YEARS",
                        required=False,
                        dest='term_in_years')
    parser.add_argument('-sp', '--run-savings-plans',
                        help="Run the savings plans recommendations True|False",
                        default=True,
                        required=False,
                        dest='run_sp')
    parser.add_argument('-ri', '--run-reserved-instances',
                        help="Run the reserved instances recommendations True|False",
                        default=True,
                        required=False,
                        dest='run_ri')
    args = parser.parse_args()

    file_name_prefix = 'CostSavingRecommendations'
    xl = ExcelSheet(file_name_prefix)
    if args.run_sp:
        write_sp_to_excel(xl, terms=args.term_in_years.split(','), payment_options=args.payment_options.split(','))
    if args.run_ri:
        write_ri_to_excel(xl, terms=args.term_in_years.split(','), payment_options=args.payment_options.split(','))
    xl.close()
