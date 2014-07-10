# -*- coding: utf-8 -*-

require 'json'
require 'turbotlib'
require 'nokogiri'
require 'faraday'
require 'roo'

def self.get_xls_links(doc, cell)
  table = doc.search('table')[1]
  list = []
  table.search('tr').each do |tr|
    td = tr.search('td')[cell]
    links = td.search('a').map { |link| "http://www.mass.gov" + link['href'] }
    list += links.select { |link| link.include?('.xls') }.uniq
  end
  list
end

Turbotlib.log("Starting run...") # optional debug logging

body = Faraday.get('http://www.mass.gov/ocabr/licensee/license-types/insurance/individual-and-business-entity-licensing/licensed-listings.html').body
doc = Nokogiri::HTML(body)

#cell 1 = agencies, cell 2 = individuals
agency_xls_links = get_xls_links(doc, 1)
individuals_xls_links = get_xls_links(doc, 2)

class IndividualRowParser
  def self.parse(row, xls_link)
    {
        number: row[0].to_i,
        license: row[1],
        licensure: row[3],
        individual: row[5],
        adress: row[7],
        city: row[8],
        state: row[9],
        zip: row[10],
        phone: row[12],
        email: row[14],
        source_url: xls_link,
        sample_date: Time.now
    }
  end
end

class AgencyRowParser
  def self.parse(row, xls_link)
    {
        number: row[0].to_i,
        license: row[1],
        licensure: row[4],
        agency: row[6],
        adress: row[8],
        city: row[9],
        state: row[11],
        zip: row[12],
        phone: row[16],
        source_url: xls_link,
        sample_date: Time.now
    }
  end
end

def parse_xls2(xls_link, parser)
  xls = Roo::Excel.new(xls_link)
  (xls.first_row..xls.last_row).each do |row_number|
    row = xls.row(row_number)
    puts JSON.dump(parser.parse(row,xls_link)) if row[0]
  end
end

agency_xls_links.each do |xls_link|
  parse_xls2(xls_link, AgencyRowParser)
end

individuals_xls_links.each do |xls_link|
  parse_xls2(xls_link, IndividualRowParser)
end